"""
Walidacja wyodrębnionych pól faktury.

Sprawdza poprawność i spójność danych - to, co doświadczony księgowy
zweryfikowałby ręcznie:
- suma kontrolna NIP (polski algorytm wag),
- spójność kwot: netto + VAT = brutto,
- logika dat: termin płatności nie wcześniej niż data wystawienia.

Zwraca listę ostrzeżeń per pole, które frontend podświetla.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any


# wagi do sumy kontrolnej NIP
_NIP_WEIGHTS = [6, 5, 7, 2, 3, 4, 5, 6, 7]


def validate_nip(nip: str | None) -> bool | None:
    """
    Sprawdza sumę kontrolną polskiego NIP.
    Zwraca True/False, albo None gdy brak danych do sprawdzenia.
    """
    if not nip:
        return None
    digits = re.sub(r"[^\d]", "", str(nip))
    if len(digits) != 10:
        return False
    checksum = sum(int(digits[i]) * _NIP_WEIGHTS[i] for i in range(9)) % 11
    if checksum == 10:
        return False
    return checksum == int(digits[9])


def _to_float(v: Any) -> float | None:
    if v is None:
        return None
    try:
        return float(str(v).replace(" ", "").replace(",", "."))
    except (ValueError, TypeError):
        return None


def _to_date(v: Any) -> datetime | None:
    if not v:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(v).strip(), fmt)
        except ValueError:
            continue
    return None


def validate_fields(fields: dict[str, Any], line_items: list | None = None) -> dict[str, str]:
    """
    Zwraca słownik {nazwa_pola: KOD_ostrzeżenia} dla pól, które nie
    przechodzą walidacji. Kody tłumaczy frontend (i18n). Puste = OK.
    Kody: nip_checksum, amounts_mismatch, net_gt_gross, due_before_issue,
          items_sum_mismatch.
    """
    warnings: dict[str, str] = {}

    for key in ("seller_nip", "buyer_nip"):
        nip = fields.get(key)
        if nip and validate_nip(nip) is False:
            warnings[key] = "nip_checksum"

    net = _to_float(fields.get("net_amount"))
    tax = _to_float(fields.get("tax_amount"))
    total = _to_float(fields.get("total_amount"))
    if net is not None and tax is not None and total is not None:
        if abs((net + tax) - total) > 0.02:
            warnings["total_amount"] = "amounts_mismatch"
    elif net is not None and total is not None and tax is None:
        if net - total > 0.02:
            warnings["total_amount"] = "net_gt_gross"

    # Kontrola krzyżowa z pozycjami. Wartość pozycji bywa NETTO (częściej)
    # albo BRUTTO - zależnie od faktury. Ostrzegamy tylko, gdy suma pozycji
    # nie pasuje do ŻADNEJ z kwot zbiorczych (netto ani brutto), z tolerancją
    # 1% na zaokrąglenia. To eliminuje fałszywe alarmy na poprawnych fakturach.
    if line_items:
        items_sum = 0.0
        any_total = False
        for it in line_items:
            v = _to_float(it.get("total") if isinstance(it, dict) else getattr(it, "total", None))
            if v is not None:
                items_sum += v
                any_total = True
        if any_total:
            targets = [x for x in (net, total) if x is not None]
            # dopasowanie z tolerancją 1% (lub 0,02 dla małych kwot)
            def _matches(target):
                tol = max(0.02, abs(target) * 0.01)
                return abs(items_sum - target) <= tol
            if targets and not any(_matches(x) for x in targets):
                warnings.setdefault("total_amount", "items_sum_mismatch")

    issue = _to_date(fields.get("issue_date"))
    due = _to_date(fields.get("due_date"))
    if issue and due and due < issue:
        warnings["due_date"] = "due_before_issue"

    return warnings


if __name__ == "__main__":
    # szybki test
    print("NIP 5842816354:", validate_nip("5842816354"))  # zależy od cyfr
    print("NIP 1234567890:", validate_nip("1234567890"))
    test = {"net_amount": "1000", "tax_amount": "230", "total_amount": "1230"}
    print("kwoty spójne:", validate_fields(test))
    test2 = {"net_amount": "1000", "tax_amount": "230", "total_amount": "9999"}
    print("kwoty niespójne:", validate_fields(test2))
