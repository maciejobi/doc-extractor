"""
Eksport danych faktury do płaskiego pliku CSV pod Power BI / Excel Power Query.

Format "tidy data": każdy wiersz to JEDNA pozycja faktury, z powtórzonymi
danymi nagłówka (numer, data, kontrahent...). To denormalizacja pod analitykę -
Power BI wczytuje taki plik wprost i buduje na nim wizualizacje.

Gdy zgromadzisz wiele faktur, wystarczy sklejić te pliki (lub wskazać folder
w Power BI) i masz gotowy zbiór do dashboardu: wydatki w czasie, wg kontrahenta,
struktura VAT itd.

Kolumny są stałe i w języku neutralnym (angielskie nazwy techniczne), żeby
model danych w Power BI był stabilny niezależnie od języka interfejsu.
"""

from __future__ import annotations

import csv
import io
from typing import Any

# stały, stabilny zestaw kolumn (kolejność ma znaczenie dla Power BI)
COLUMNS = [
    "invoice_number", "issue_date", "service_date", "due_date",
    "seller_name", "seller_nip", "buyer_name", "buyer_nip",
    "currency", "payment_method",
    "line_no", "description", "quantity", "unit_price", "vat_rate", "line_total",
    "invoice_net", "invoice_tax", "invoice_gross",
    "source_file",
]

_ALIASES = {"seller_name": "seller", "buyer_name": "buyer"}


def _get(fields: dict, key: str):
    if fields.get(key) not in (None, "", []):
        return fields[key]
    alias = _ALIASES.get(key)
    if alias and fields.get(alias) not in (None, "", []):
        return fields[alias]
    return ""


def _num(v: Any):
    if v is None or v == "":
        return ""
    try:
        return float(str(v).replace(" ", "").replace(",", "."))
    except (ValueError, TypeError):
        return ""


def build_powerbi_csv(data: dict[str, Any]) -> bytes:
    """Buduje płaski CSV (UTF-8 z BOM, żeby Excel/Power BI poprawnie czytał PL znaki)."""
    fields = data.get("fields", {}) or {}
    items = data.get("line_items", []) or []
    source = data.get("source_file", "")

    # wspólne dane nagłówka (powtarzane w każdym wierszu)
    header = {
        "invoice_number": _get(fields, "invoice_number"),
        "issue_date": _get(fields, "issue_date"),
        "service_date": _get(fields, "service_date"),
        "due_date": _get(fields, "due_date"),
        "seller_name": _get(fields, "seller_name"),
        "seller_nip": _get(fields, "seller_nip"),
        "buyer_name": _get(fields, "buyer_name"),
        "buyer_nip": _get(fields, "buyer_nip"),
        "currency": _get(fields, "currency"),
        "payment_method": _get(fields, "payment_method"),
        "invoice_net": _num(_get(fields, "net_amount")),
        "invoice_tax": _num(_get(fields, "tax_amount")),
        "invoice_gross": _num(_get(fields, "total_amount")),
        "source_file": source,
    }

    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=COLUMNS, extrasaction="ignore")
    writer.writeheader()

    if items:
        for i, it in enumerate(items, start=1):
            if not isinstance(it, dict):
                continue
            row = dict(header)
            row.update({
                "line_no": i,
                "description": it.get("description", ""),
                "quantity": _num(it.get("quantity")),
                "unit_price": _num(it.get("unit_price")),
                "vat_rate": it.get("vat_rate", "") or "",
                "line_total": _num(it.get("total")),
            })
            writer.writerow(row)
    else:
        # faktura bez rozpoznanych pozycji - jeden wiersz z samym nagłówkiem
        row = dict(header)
        row.update({"line_no": "", "description": "", "quantity": "",
                    "unit_price": "", "vat_rate": "", "line_total": ""})
        writer.writerow(row)

    # BOM dla poprawnego odczytu polskich znaków w Excelu
    return "\ufeff".encode("utf-8") + out.getvalue().encode("utf-8")


if __name__ == "__main__":
    demo = {
        "source_file": "fv.pdf",
        "fields": {"invoice_number": "FV/1", "issue_date": "2026-01-15",
                   "seller": "Test", "currency": "PLN", "total_amount": "2676"},
        "line_items": [
            {"description": "A", "quantity": "2", "unit_price": "1000", "vat_rate": "23%", "total": "2460"},
            {"description": "B", "quantity": "1", "unit_price": "200", "vat_rate": "8%", "total": "216"},
        ],
    }
    print(build_powerbi_csv(demo).decode("utf-8"))
