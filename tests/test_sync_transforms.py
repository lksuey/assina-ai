"""Tests for the retention and call-log rules in sync/sync_to_sheets.py.

Run: python -m unittest discover -s tests -p "test_*.py"   (pytest also works if you have it)
Needs pandas (see sync/requirements.txt); gspread is NOT needed because it is imported lazily.
"""
import importlib.util
import sys
import unittest
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    pd = None

SPEC_PATH = Path(__file__).resolve().parents[1] / "sync" / "sync_to_sheets.py"
NOW = datetime(2026, 10, 30, 15, 0, 0)


def load_module():
    spec = importlib.util.spec_from_file_location("sync_to_sheets", SPEC_PATH)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["sync_to_sheets"] = mod
    spec.loader.exec_module(mod)
    return mod


@unittest.skipIf(pd is None, "pandas not installed")
class RetentionRules(unittest.TestCase):
    def setUp(self):
        self.m = load_module()

    def frame(self):
        return pd.DataFrame([
            {"N° Proposta": "1", "Status Admin.": "Assinado", "Alt. Status": "", "Emissão": "", "Solic. Proposta": "2026-01-02 10:00:00"},   # active and very old
            {"N° Proposta": "2", "Status Admin.": "Pago", "Alt. Status": "2026-10-25", "Emissão": "", "Solic. Proposta": "2026-10-20 10:00:00"},  # terminal, recent
            {"N° Proposta": "3", "Status Admin.": "Cancelado", "Alt. Status": "2026-08-01", "Emissão": "", "Solic. Proposta": "2026-07-30 10:00:00"},  # terminal, old
            {"N° Proposta": "4", "Status Admin.": "cancelado ", "Alt. Status": "", "Emissão": "", "Solic. Proposta": "2026-10-28 09:00:00"},  # fallback to Solic. date, case/space tolerant
        ])

    def test_active_rows_are_never_cut_and_terminal_rows_follow_the_window(self):
        out = self.m._filtrar_janela_ativa(self.frame(), 45, agora=NOW)
        self.assertEqual(sorted(out["N° Proposta"]), ["1", "2", "4"])

    def test_reference_date_falls_back_in_order(self):
        row = pd.Series({"Alt. Status": "", "Emissão": "30/09/2026", "Solic. Proposta": "2026-01-01 00:00:00"})
        self.assertEqual(self.m._data_referencia(row), pd.Timestamp("2026-09-30"))
        self.assertIsNone(self.m._data_referencia(pd.Series({"Alt. Status": "", "Emissão": "", "Solic. Proposta": ""})))

    def test_api_cutoff_only_touches_api_rows(self):
        self.m.DATA_CORTE_API = "2026-06-01"
        df = pd.DataFrame([
            {"Analista": "API CONSIGNADO APP", "Solic. Proposta": "2026-05-30 10:00:00"},
            {"Analista": "API CONSIGNADO APP", "Solic. Proposta": "2026-06-02 10:00:00"},
            {"Analista": "Analyst 01", "Solic. Proposta": "2026-05-30 10:00:00"},
        ])
        out = self.m._filtrar_retencao_api(df)
        self.assertEqual(len(out), 2)
        self.assertEqual(sorted(out["Analista"]), ["API CONSIGNADO APP", "Analyst 01"])

    def test_filter_is_a_noop_without_the_status_column(self):
        df = pd.DataFrame([{"x": 1}])
        self.assertEqual(len(self.m._filtrar_janela_ativa(df, 10, agora=NOW)), 1)


@unittest.skipIf(pd is None, "pandas not installed")
class CallLog(unittest.TestCase):
    def setUp(self):
        self.m = load_module()

    def test_duration_parser(self):
        f = self.m._duracao_para_segundos
        self.assertEqual(f("00:01:30"), 90)
        self.assertEqual(f("01:00:00"), 3600)
        self.assertEqual(f("bad"), 0)
        self.assertEqual(f(None), 0)
        self.assertEqual(f("1:2"), 0)

    def test_timestamp_formats_are_tried_in_chain(self):
        s = pd.Series(["28/10/2026 11:54", "Oct 28, 2026 11:54 AM", "2026-10-28 11:54:00", "garbage"])
        out = self.m.parse_datas_call_log(s)
        self.assertEqual(out.iloc[0], pd.Timestamp("2026-10-28 11:54"))
        self.assertEqual(out.iloc[1], pd.Timestamp("2026-10-28 11:54"))
        self.assertEqual(out.iloc[2], pd.Timestamp("2026-10-28 11:54"))
        self.assertTrue(pd.isna(out.iloc[3]))

    def test_prepare_keeps_only_recent_outbound_calls_with_fallback_duration(self):
        d = self.m.DIRECAO_SAIDA
        df = pd.DataFrame([
            {"Destino": "(00) 90000-0001", "Direcao": d, "Data_Inicial_Chamada": "2026-10-29 10:00:00", "Tempo_Conversa": "00:00:40", "Tempo_Atendimento": None},
            {"Destino": "00900000002", "Direcao": d, "Data_Inicial_Chamada": "2026-10-29 11:00:00", "Tempo_Conversa": None, "Tempo_Atendimento": "00:00:05"},
            {"Destino": "00900000003", "Direcao": "Inbound", "Data_Inicial_Chamada": "2026-10-29 11:00:00", "Tempo_Conversa": "00:01:00", "Tempo_Atendimento": None},
            {"Destino": "00900000004", "Direcao": d, "Data_Inicial_Chamada": "2020-01-01 11:00:00", "Tempo_Conversa": "00:01:00", "Tempo_Atendimento": None},
            {"Destino": "", "Direcao": d, "Data_Inicial_Chamada": "2026-10-29 11:00:00", "Tempo_Conversa": "00:01:00", "Tempo_Atendimento": None},
        ])
        out = self.m.preparar_call_log(df, agora=NOW)
        self.assertEqual(list(out.columns), ["Celular", "DataHora", "Segundos"])
        self.assertEqual(out["Celular"].tolist(), ["00900000001", "00900000002"])
        self.assertEqual(out["Segundos"].tolist(), [40, 5])
        self.assertEqual(out["DataHora"].iloc[0], "29/10/2026 10:00:00")


if __name__ == "__main__":
    unittest.main()
