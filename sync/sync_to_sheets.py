"""Sync local spreadsheet exports to the Google Sheet that feeds the Apps Script dashboard.

The dashboard reads from a Google Sheet (tabs "Raw", "TotalPropostas", "Efetivados",
"EfetivadosAPI", "Supervisores", "CallLog", "Config"). This script pushes local
.xlsx exports into those tabs, applying retention rules so the sheet stays small
and the dashboard stays fast.

Configuration is entirely through environment variables (nothing is hard coded):

    GOOGLE_CREDENTIALS_JSON   path to a service account JSON key (never commit it)
    SPREADSHEET_ID            id of the target Google Sheet
    RAW_XLSX                  active pipeline export (feeds the "Raw" tab)
    TOTAL_XLSX                all proposals export (feeds "TotalPropostas")
    PAID_XLSX                 paid proposals export (feeds "Efetivados")
    PAID_API_XLSX             paid proposals of the API channels (feeds "EfetivadosAPI")
    SUPERVISORS_XLSX          analyst to supervisor mapping
    CALL_LOG_XLSX             outbound call log

Optional: RETENTION_DAYS_RAW (45), RETENTION_DAYS (180), DATA_CORTE_API (fixed cutoff
date for API channel rows, default 2000-01-01).

One-time setup, outside this script: create a Google Cloud project, enable the Sheets
API, create a service account, share the target sheet with the service account
e-mail (Editor), then ``pip install -r requirements.txt``.
"""
import os
import re
import sys
import time
from datetime import datetime, timedelta

import pandas as pd

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

CREDENTIALS_PATH = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")

RAW_XLSX = os.environ.get("RAW_XLSX", "")
TOTAL_XLSX = os.environ.get("TOTAL_XLSX", "")
PAID_XLSX = os.environ.get("PAID_XLSX", "")
PAID_API_XLSX = os.environ.get("PAID_API_XLSX", "")
SUPERVISORS_XLSX = os.environ.get("SUPERVISORS_XLSX", "")
CALL_LOG_XLSX = os.environ.get("CALL_LOG_XLSX", "")

# Only the columns the management cards use; the source exports have dozens of
# columns and tens of thousands of rows. "N° Proposta" is the join key between
# Raw and the paid tabs (NOT "ID Cliente": one customer can have many proposals
# and only one of them may have been paid, so joining on the customer would mark
# every proposal of that customer as converted).
COLUNAS_GERENCIAL = ["Solic. Proposta", "Hora Solic. Proposta", "Val. Solic.", "Seguro",
                     "Analista", "Status Admin.", "Correspondente", "ID Cliente", "N° Proposta"]
ABAS_GERENCIAL = [
    ("TOTAL_XLSX", "TotalPropostas"),
    ("PAID_XLSX", "Efetivados"),
    ("PAID_API_XLSX", "EfetivadosAPI"),
]

# Only outbound calls count as "customer contact attempts".
DIRECAO_SAIDA = os.environ.get("DIRECAO_SAIDA", "Outbound")
# Calls shorter than this (talk time, falling back to answer time) are treated as
# technical drops or voicemail, not a real contact.
DURACAO_MINIMA_SUCESSO_SEGUNDOS = 3

# The local history only grows; the Raw tab must not mirror it forever. Only
# proposals still in the active pipeline, or that became Paid/Cancelled within the
# window, are synced. Active proposals are never cut by the window.
RETENTION_DAYS_RAW = int(os.environ.get("RETENTION_DAYS_RAW", "45"))
RETENTION_DAYS = int(os.environ.get("RETENTION_DAYS", "180"))
STATUS_COL = "Status Admin."
STATUS_TERMINAIS = {"pago", "cancelado"}
COLUNAS_DATA_REFERENCIA = ["Alt. Status", "Emissão", "Solic. Proposta"]

# API channel rows have their own fixed (not rolling) cutoff.
DATA_CORTE_API = os.environ.get("DATA_CORTE_API", "2000-01-01")
ANALISTA_COL = "Analista"


def _eh_api(df: pd.DataFrame) -> pd.Series:
    if ANALISTA_COL not in df.columns:
        return pd.Series(False, index=df.index)
    return df[ANALISTA_COL].astype(str).str.strip().str.upper().str.startswith("API ")


def _filtrar_retencao_api(df: pd.DataFrame, col_data: str = "Solic. Proposta") -> pd.DataFrame:
    """Apply the fixed API cutoff to API channel rows only; call center rows pass through."""
    if df.empty or col_data not in df.columns:
        return df
    api = _eh_api(df)
    if not api.any():
        return df
    corte = pd.to_datetime(DATA_CORTE_API)
    datas = pd.to_datetime(df[col_data].astype(str).str[:10], format="%Y-%m-%d", errors="coerce")
    manter = (~api) | (datas >= corte)
    descartadas = int((api & ~manter).sum())
    if descartadas:
        print(f"[info] {descartadas} API rows before {DATA_CORTE_API} were not synced "
              f"(they stay in the local history).")
    return df[manter]


def _client():
    import gspread  # imported lazily so the pure functions above work without it
    from google.oauth2.service_account import Credentials

    if not CREDENTIALS_PATH or not os.path.exists(CREDENTIALS_PATH):
        sys.exit("Set GOOGLE_CREDENTIALS_JSON to the path of the service account key (see the module docstring).")
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    return gspread.authorize(creds)


def _worksheet_ou_criar(gc, titulo: str):
    """Create the tab on the fly so a new tab never blocks the sync."""
    import gspread

    ss = gc.open_by_key(SPREADSHEET_ID)
    try:
        return ss.worksheet(titulo)
    except gspread.exceptions.WorksheetNotFound:
        print(f"[info] tab \"{titulo}\" did not exist, creating it.")
        return ss.add_worksheet(title=titulo, rows=1000, cols=50)


def _write_dataframe(ws, df: pd.DataFrame, tentativas: int = 3, espera_s: int = 15):
    """Write the DataFrame from A1 and only clear leftover old rows AFTER the new
    write is confirmed.

    A previous version called ws.clear() and ws.update() as two separate API calls;
    when the second one failed midway (network, timeout, quota) the tab stayed
    empty until the next sync. Writing first and trimming afterwards, with retries,
    means a network failure never leaves the tab empty."""
    df = df.fillna("")
    linhas = [df.columns.tolist()] + df.astype(str).values.tolist()
    linhas_antigas = ws.row_count

    ultimo_erro = None
    for tentativa in range(tentativas):
        try:
            for inicio in range(0, len(linhas), 10000):
                ws.update(linhas[inicio:inicio + 10000], range_name=f"A{inicio + 1}")
            if linhas_antigas > len(linhas):
                ws.batch_clear([f"A{len(linhas) + 1}:ZZ{linhas_antigas}"])
            return
        except Exception as e:  # noqa: BLE001 - retry on any transient failure
            ultimo_erro = e
            print(f"  [retry {tentativa + 1}/{tentativas}] write failed: {e}")
            if tentativa < tentativas - 1:
                time.sleep(espera_s)
    raise ultimo_erro


def _data_referencia(row: pd.Series):
    for col in COLUNAS_DATA_REFERENCIA:
        valor = row.get(col)
        if not valor or not str(valor).strip():
            continue
        bruto = str(valor)[:10]
        fmt = "%Y-%m-%d" if re.match(r"^\d{4}-\d{2}-\d{2}$", bruto) else None
        dt = pd.to_datetime(bruto, format=fmt, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt
    return None


def _filtrar_janela_ativa(df: pd.DataFrame, dias: int, agora: datetime = None) -> pd.DataFrame:
    """Keep every proposal still in the pipeline, plus terminal ones inside the window."""
    if STATUS_COL not in df.columns or df.empty:
        return df
    limite = (agora or datetime.now()) - timedelta(days=dias)
    status_norm = df[STATUS_COL].astype(str).str.strip().str.lower()
    ativas = ~status_norm.isin(STATUS_TERMINAIS)

    datas_ref = df.apply(_data_referencia, axis=1)
    dentro_da_janela = datas_ref.apply(lambda d: d is not None and d >= limite)

    manter = ativas | dentro_da_janela
    descartadas = int((~manter).sum())
    if descartadas:
        print(f"[info] {descartadas} terminal proposals outside the {dias}-day window were not synced.")
    return df[manter]


def sync_raw(gc) -> None:
    if not RAW_XLSX or not os.path.exists(RAW_XLSX):
        print("[warn] RAW_XLSX not set or not found, skipping Raw.")
        return
    df = pd.read_excel(RAW_XLSX, dtype=str, engine="openpyxl")
    total_local = len(df)
    df = _filtrar_janela_ativa(df, RETENTION_DAYS_RAW)
    df = _filtrar_retencao_api(df)
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet("Raw")
    _write_dataframe(ws, df)
    print(f"Raw synced: {len(df)} of {total_local} local rows (active window {RETENTION_DAYS_RAW} days).")


def sync_gerencial(gc) -> None:
    """TOTAL/PAID/PAID_API exports -> TotalPropostas/Efetivados/EfetivadosAPI.

    Cut by "Solic. Proposta" inside RETENTION_DAYS (the cards count by that date).
    EfetivadosAPI uses the fixed DATA_CORTE_API instead of the rolling window."""
    limite = datetime.now() - timedelta(days=RETENTION_DAYS)
    corte_api = pd.to_datetime(DATA_CORTE_API)
    for env_name, aba in ABAS_GERENCIAL:
        caminho = os.environ.get(env_name, "")
        if not caminho or not os.path.exists(caminho):
            print(f"[warn] {env_name} not set or not found, skipping {aba}.")
            continue
        df = pd.read_excel(caminho, dtype=str, engine="openpyxl", usecols=COLUNAS_GERENCIAL)
        total = len(df)
        datas = pd.to_datetime(df["Solic. Proposta"].str[:10], format="%Y-%m-%d", errors="coerce")
        limite_aba = corte_api if aba == "EfetivadosAPI" else limite
        df = df[datas >= limite_aba].fillna("")

        # read_excel(usecols=...) returns columns in the FILE order, not the order of
        # the usecols list. Always use df.columns as the header, never the manual
        # list, or header and values drift apart.
        colunas_reais = df.columns.tolist()
        ws = _worksheet_ou_criar(gc, aba)
        ws.clear()
        ws.resize(rows=len(df) + 1, cols=len(colunas_reais))
        linhas = [colunas_reais] + df.astype(str).values.tolist()
        # In chunks: tens of thousands of rows in one request exceed the API payload limit.
        for inicio in range(0, len(linhas), 10000):
            ws.update(linhas[inicio:inicio + 10000], range_name=f"A{inicio + 1}")
        print(f"{aba} synced: {len(df)} of {total} rows (last {RETENTION_DAYS} days).")


def sync_supervisores(gc) -> None:
    if not SUPERVISORS_XLSX or not os.path.exists(SUPERVISORS_XLSX):
        print("[warn] SUPERVISORS_XLSX not set or not found, skipping Supervisores.")
        return
    df = pd.read_excel(SUPERVISORS_XLSX, dtype=str, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    analista_col = next((c for c in df.columns if "analista" in c.lower()), None)
    supervisor_col = next((c for c in df.columns if "supervisor" in c.lower()), None)
    if not analista_col or not supervisor_col:
        print("[warn] could not find Analista/Supervisor columns, skipping.")
        return
    df = df[[analista_col, supervisor_col]].rename(columns={analista_col: "Analista", supervisor_col: "Supervisor"})
    df = df.dropna(subset=["Analista"])
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet("Supervisores")
    _write_dataframe(ws, df)
    print(f"Supervisores synced: {len(df)} rows.")


def _duracao_para_segundos(hhmmss) -> int:
    if not hhmmss or not isinstance(hhmmss, str):
        return 0
    partes = hhmmss.split(":")
    if len(partes) != 3:
        return 0
    try:
        h, m, s = (int(p) for p in partes)
        return h * 3600 + m * 60 + s
    except ValueError:
        return 0


# The source has changed its timestamp format more than once without notice. Each
# known format is tried in order and, as a last resort, pandas infers it. Without
# this chain a format change silently turned every row into NaT and the "returned
# to customer" flag was always "No".
FORMATOS_DATA_CONHECIDOS = ["%d/%m/%Y %H:%M", "%b %d, %Y %I:%M %p", "%Y-%m-%d %H:%M:%S"]


def parse_datas_call_log(serie: pd.Series) -> pd.Series:
    resultado = pd.Series(pd.NaT, index=serie.index)
    for fmt in FORMATOS_DATA_CONHECIDOS:
        falta = resultado.isna()
        if not falta.any():
            break
        resultado.loc[falta] = pd.to_datetime(serie.loc[falta], format=fmt, errors="coerce")
    falta = resultado.isna()
    if falta.any():
        resultado.loc[falta] = pd.to_datetime(serie.loc[falta], errors="coerce")
    return resultado


def preparar_call_log(df: pd.DataFrame, agora: datetime = None) -> pd.DataFrame:
    """Turn the raw call log into the compact (Celular, DataHora, Segundos) shape.

    One row per call, not pre-aggregated per phone: the dashboard compares each call
    against the timestamp of EACH proposal (only calls after the proposal count)."""
    df = df[df["Direcao"] == DIRECAO_SAIDA].copy()
    df["Celular"] = df["Destino"].astype(str).str.replace(r"\D", "", regex=True)
    df = df[df["Celular"] != ""]

    df["_dt"] = parse_datas_call_log(df["Data_Inicial_Chamada"])
    ainda_falta = int(df["_dt"].isna().sum())
    if ainda_falta:
        print(f"[warn] {ainda_falta} call log rows with an unrecognised timestamp format were dropped.")
    limite = (agora or datetime.now()) - timedelta(days=RETENTION_DAYS)
    df = df[df["_dt"].notna() & (df["_dt"] >= limite)]

    df["Segundos"] = df["Tempo_Conversa"].fillna(df["Tempo_Atendimento"]).apply(_duracao_para_segundos)
    df["DataHora"] = df["_dt"].dt.strftime("%d/%m/%Y %H:%M:%S")
    return df[["Celular", "DataHora", "Segundos"]].reset_index(drop=True)


def sync_call_log(gc) -> None:
    if not CALL_LOG_XLSX or not os.path.exists(CALL_LOG_XLSX):
        print("[warn] CALL_LOG_XLSX not set or not found, skipping CallLog.")
        return
    df = pd.read_excel(
        CALL_LOG_XLSX, dtype=str, engine="openpyxl",
        usecols=["Destino", "Direcao", "Data_Inicial_Chamada", "Tempo_Conversa", "Tempo_Atendimento"],
    )
    saida = preparar_call_log(df)
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet("CallLog")
    _write_dataframe(ws, saida)
    print(f"CallLog synced: {len(saida)} outbound calls (window {RETENTION_DAYS} days).")


def _atualizar_ultima_sincronizacao(gc) -> None:
    """Record when the sync really ran in the Config tab: that value (not the
    server response time) drives the freshness indicator in the dashboard."""
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet("Config")
    valores = ws.get_all_values()
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    for i, linha in enumerate(valores[1:], start=2):
        if linha and linha[0] == "Ultima_Sincronizacao":
            ws.update_cell(i, 2, agora)
            return
    ws.append_row(["Ultima_Sincronizacao", agora])


def sincronizar_tudo() -> None:
    if not SPREADSHEET_ID:
        sys.exit("Set SPREADSHEET_ID (environment variable).")
    gc = _client()
    sync_raw(gc)
    sync_gerencial(gc)
    sync_supervisores(gc)
    sync_call_log(gc)
    _atualizar_ultima_sincronizacao(gc)
    print("Sync complete.")


def main() -> None:
    sincronizar_tudo()


if __name__ == "__main__":
    main()
