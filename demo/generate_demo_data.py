"""Generate fully synthetic demo data for the Assina AI dashboard.

Everything here is invented: customers are "Customer 001", analysts are "Analyst 01",
correspondents are "Correspondent A/B", phones use the invalid area code 00 and ids are
sequential. The output is deterministic for a given (seed, anchor date).

    python demo/generate_demo_data.py                    # anchor = today, writes demo/demo_data.json
    python demo/generate_demo_data.py --anchor 2026-10-30 --out demo/demo_data.json
"""
import argparse
import json
import random
from datetime import date, datetime, timedelta
from pathlib import Path

SEED = 20260101
DAYS = 35
N_CUSTOMERS = 520

CORRESPONDENT_A = "CORRESPONDENT A - 0001"
CORRESPONDENT_B = "CORRESPONDENT B - 0002"
CORRESPONDENT_UNKNOWN = "CORRESPONDENT Z - 0099"

API_ANALYSTS = [
    "API CONSIGNADO CHANNEL A", "API CONSIGNADO CHANNEL B", "API CONSIGNADO APP",
    "API CONSIGNADO LANDING PAGE", "API CONSIGNADO LEADER", "API CONSIGNADO PARTNER DESK",
]
CALL_CENTER_ANALYSTS = ["Analyst %02d" % i for i in range(1, 13)]
SUPERVISORS = {a: "Supervisor %s" % "ABC"[i % 3] for i, a in enumerate(CALL_CENTER_ANALYSTS)}

STATUS_WEIGHTS = [
    ("Pago", 22), ("Cancelado", 18), ("Assinado o primeiro", 9), ("Não assinado", 12),
    ("Assinado", 7), ("Conferência", 6), ("Conferência de fraude", 3), ("Pendente regularização", 4),
    ("Processando pagamento automático", 3), ("Banco", 5), ("Liberado", 5), ("Pendente", 6),
]
CANCEL_REASONS = [
    "Customer gave up", "Biometric validation failed", "Document mismatch", "Payroll margin exceeded",
    "REGISTRY RETURN FAILURE: CET RATE 1.00%",
    "REGISTRY RETURN FAILURE: TN - Monthly interest rate is invalid: it must be lower than or equal to the CET rate, with a spread of at most 1.00%",
    "",
    "#N/D",
]

RAW_HEADERS = ["Solic. Proposta", "Hora Solic. Proposta", "Alt. Status", "N° Proposta", "ID Cliente", "Cliente",
               "Celular", "Val. Solic.", "Seguro", "Analista", "Status Admin.", "Correspondente", "Motivo_Cancelamento"]


def money(rng):
    value = rng.choice([1500, 2300, 3100, 4800, 6200, 8900, 12000, 15500]) + rng.randint(0, 99)
    cents = rng.randint(0, 99)
    return "{:,.2f}".format(value + cents / 100).replace(",", "X").replace(".", ",").replace("X", ".")


def phone(customer_index, variant):
    base = "(00) 90000-%04d" % customer_index
    return base if variant == 0 else "00900000%04d" % customer_index


def build(anchor: date, seed: int = SEED):
    rng = random.Random(seed)
    customers = [{"id": "C%05d" % i, "name": "Customer %03d" % i, "idx": i} for i in range(1, N_CUSTOMERS + 1)]
    statuses, weights = zip(*STATUS_WEIGHTS)

    raw, total, paid, paid_api = [], [], [], []
    calls = []
    proposal_no = 100000
    for day_offset in range(DAYS, -1, -1):
        day = anchor - timedelta(days=day_offset)
        if day.weekday() == 6:
            continue
        volume = rng.randint(22, 48) if day.weekday() < 5 else rng.randint(6, 14)
        if day_offset == 0:
            volume = rng.randint(10, 20)
        for _ in range(volume):
            proposal_no += 1
            customer = rng.choice(customers)
            is_api = rng.random() < 0.38
            if is_api:
                analyst = rng.choice(API_ANALYSTS)
                correspondent = CORRESPONDENT_A
            else:
                analyst = rng.choice(CALL_CENTER_ANALYSTS)
                correspondent = CORRESPONDENT_B
            status = rng.choices(statuses, weights)[0]
            moment = datetime(day.year, day.month, day.day, rng.randint(8, 19), rng.randint(0, 59), rng.randint(0, 59))
            if day_offset == 0:
                moment = moment.replace(hour=rng.randint(8, 11))
            reason = rng.choice(CANCEL_REASONS) if status == "Cancelado" else ""
            row = [
                moment.strftime("%Y-%m-%d %H:%M:%S"), moment.strftime("%H:%M:%S"),
                moment.strftime("%Y-%m-%d") if rng.random() < 0.6 else "",
                str(proposal_no), customer["id"], customer["name"], phone(customer["idx"], rng.randint(0, 1)),
                money(rng), rng.choice(["0,00", "0,00", "189,90", "249,00"]),
                analyst, status, correspondent, reason,
            ]
            raw.append(row)
            if status == "Pago":
                (paid_api if is_api else paid).append(row)
            if not is_api:
                total.append(row)
            if status in ("Assinado o primeiro", "Não assinado", "Cancelado", "Assinado") and rng.random() < 0.55:
                for _ in range(rng.randint(1, 3)):
                    call_at = moment + timedelta(minutes=rng.randint(5, 600))
                    calls.append([phone(customer["idx"], 0), call_at.strftime("%d/%m/%Y %H:%M:%S"),
                                  rng.choice([0, 1, 2, 4, 35, 120, 300])])

    for k in range(6):
        proposal_no += 1
        c = customers[k]
        moment = datetime(anchor.year, anchor.month, anchor.day, 9, k, 0) - timedelta(days=1 + k % 3)
        bad_corr = CORRESPONDENT_UNKNOWN if k % 2 == 0 else CORRESPONDENT_A
        raw.append([moment.strftime("%Y-%m-%d %H:%M:%S"), moment.strftime("%H:%M:%S"), "", str(proposal_no),
                    "X%05d" % c["idx"], "Customer %03d" % c["idx"], phone(c["idx"], 0), "1.000,00", "0,00",
                    "Analyst 01", "Assinado", bad_corr, ""])

    raw.sort(key=lambda r: (r[0], r[3]))
    total.sort(key=lambda r: (r[0], r[3]))
    paid.sort(key=lambda r: (r[0], r[3]))
    paid_api.sort(key=lambda r: (r[0], r[3]))

    sync_stamp = datetime(anchor.year, anchor.month, anchor.day, 7, 55, 0).strftime("%d/%m/%Y %H:%M:%S")
    return {
        "meta": {"seed": seed, "anchor": anchor.isoformat(), "synthetic": True},
        "sheets": {
            "Raw": {"headers": RAW_HEADERS, "rows": raw},
            "TotalPropostas": {"headers": RAW_HEADERS, "rows": total},
            "Efetivados": {"headers": RAW_HEADERS, "rows": paid},
            "EfetivadosAPI": {"headers": RAW_HEADERS, "rows": paid_api},
            "Supervisores": {"headers": ["Analista", "Supervisor"], "rows": [[a, s] for a, s in SUPERVISORS.items()]},
            "CallLog": {"headers": [], "rows": calls},
            "Config": {"headers": ["Chave", "Valor"], "rows": [
                ["Meta_Conversao_Diaria", 20], ["Limiar_Alerta_Vermelho", 0.15], ["Dias_Media_Movel", 7],
                ["Horario_Corte_Dia", "23:59"], ["Ultima_Sincronizacao", sync_stamp]]},
        },
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--anchor", default=date.today().isoformat(), help="last day of data, YYYY-MM-DD (default: today)")
    parser.add_argument("--seed", type=int, default=SEED)
    parser.add_argument("--out", default=str(Path(__file__).with_name("demo_data.json")))
    args = parser.parse_args()
    data = build(date.fromisoformat(args.anchor), args.seed)
    Path(args.out).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    sheets = data["sheets"]
    print("wrote %s: %s" % (args.out, ", ".join("%s=%d" % (k, len(v["rows"])) for k, v in sheets.items())))


if __name__ == "__main__":
    main()
