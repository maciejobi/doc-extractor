"""
Ekstrakcja pól faktury za pomocą LLM (Claude).

Kluczowa cecha: DZIAŁA WARUNKOWO.
- Jeśli ustawiona jest zmienna środowiskowa ANTHROPIC_API_KEY -> używa LLM,
  który czyta układ faktury "jak człowiek" (poprawnie przypisuje NIP-y do stron,
  rozumie tabele o dowolnym układzie, radzi sobie z różnymi formatami).
- Jeśli klucza brak -> zwraca None, a aplikacja płynnie wraca do parsera
  regułowego. Dzięki temu repo działa u każdego bez konfiguracji.

Klucz NIGDY nie jest w kodzie - tylko ze zmiennej środowiskowej.
"""

from __future__ import annotations

import json
import os
from typing import Any

# Schemat pól, które chcemy wyciągnąć. LLM dostaje go w promptcie i zwraca
# dokładnie taką strukturę - to jest "structured output" wymuszony instrukcją.
INVOICE_SCHEMA = {
    "invoice_number": "numer faktury (string lub null)",
    "issue_date": "data wystawienia w formacie YYYY-MM-DD (lub null)",
    "service_date": "data wykonania/sprzedaży YYYY-MM-DD (lub null)",
    "due_date": "termin płatności YYYY-MM-DD (lub null)",
    "seller_name": "pełna nazwa sprzedawcy (lub null)",
    "seller_nip": "NIP sprzedawcy, same cyfry (lub null)",
    "seller_address": "adres sprzedawcy jedną linią (lub null)",
    "buyer_name": "pełna nazwa nabywcy (lub null)",
    "buyer_nip": "NIP nabywcy, same cyfry (lub null)",
    "buyer_address": "adres nabywcy jedną linią (lub null)",
    "net_amount": "suma netto jako liczba (lub null)",
    "tax_amount": "suma VAT jako liczba (lub null)",
    "total_amount": "kwota brutto do zapłaty jako liczba (lub null)",
    "currency": "waluta, np. PLN (lub null)",
    "payment_method": "metoda płatności (lub null)",
    "bank_account": "numer konta / IBAN (lub null)",
    "reference_numbers": "numery referencyjne oddzielone przecinkami (lub null)",
    "line_items": [
        {
            "description": "nazwa pozycji",
            "quantity": "ilość jako liczba (lub null)",
            "unit": "jednostka miary (lub null)",
            "unit_price": "cena netto jednostkowa jako liczba (lub null)",
            "vat_rate": "stawka VAT, np. '23%' (lub null)",
            "total": "wartość pozycji jako liczba (lub null)",
        }
    ],
}


def is_available() -> bool:
    """Czy LLM jest skonfigurowany (jest klucz i biblioteka)?"""
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return False
    try:
        import anthropic  # noqa: F401
        return True
    except ImportError:
        return False


def extract_with_llm(text: str) -> dict[str, Any] | None:
    """
    Wyciąga pola faktury z surowego tekstu za pomocą Claude.
    Zwraca słownik zgodny ze schematem albo None, jeśli LLM niedostępny
    lub wystąpił błąd (wtedy woła się fallback regułowy).
    """
    if not is_available():
        return None

    try:
        import anthropic
    except ImportError:
        return None

    client = anthropic.Anthropic()  # klucz brany automatycznie z env
    model = os.environ.get("LLM_MODEL", "claude-haiku-4-5-20251001")

    schema_str = json.dumps(INVOICE_SCHEMA, indent=2, ensure_ascii=False)
    prompt = (
        "Jesteś precyzyjnym ekstraktorem danych z faktur. Poniżej tekst "
        "wyciągnięty z faktury (PDF/OCR). Wyodrębnij dane zgodnie ze schematem.\n\n"
        "ZASADY:\n"
        "- Zwróć WYŁĄCZNIE poprawny JSON, bez komentarzy i bez ```.\n"
        "- Przypisuj NIP i adres do właściwej strony (sprzedawca vs nabywca) "
        "na podstawie układu i nagłówków, nie kolejności w tekście.\n"
        "- Kwoty jako liczby (kropka dziesiętna, bez spacji i symbolu waluty).\n"
        "- Daty w formacie YYYY-MM-DD.\n"
        "- Brakujące pola ustaw na null. Nie zgaduj.\n\n"
        "ZASADY DLA KWOT ZBIORCZYCH (ważne):\n"
        "- total_amount = kwota brutto do zapłaty (suma należności, wartość 'razem').\n"
        "- net_amount = suma wartości NETTO wszystkich pozycji.\n"
        "- tax_amount = kwota podatku VAT. Jeśli faktura nie podaje jej wprost, "
        "policz: tax_amount = total_amount - net_amount. NIGDY nie kopiuj tu "
        "wartości netto pojedynczej pozycji ani stawki procentowej.\n"
        "- Kolumna 'VAT' w tabeli pozycji to zwykle STAWKA procentowa (np. 23%), "
        "a NIE kwota podatku - nie myl jej z tax_amount.\n"
        "- Sprawdź spójność: net_amount + tax_amount musi równać się total_amount.\n\n"
        f"SCHEMAT:\n{schema_str}\n\n"
        f"TEKST FAKTURY:\n{text[:8000]}"
    )

    try:
        resp = client.messages.create(
            model=model,
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = "".join(
            block.text for block in resp.content if block.type == "text"
        ).strip()
        # zabezpieczenie, gdyby model owinął w ```json
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return json.loads(raw)
    except Exception:  # noqa: BLE001 - każdy błąd -> fallback regułowy
        return None


def extract_from_image(image_bytes: bytes, media_type: str) -> dict[str, Any] | None:
    """
    Wyciąga pola faktury bezpośrednio ze ZDJĘCIA (JPG/PNG) za pomocą Claude.
    Claude czyta obraz natywnie - lepiej radzi sobie z krzywymi, telefonicznymi
    zdjęciami niż klasyczny OCR. Zwraca słownik zgodny ze schematem albo None.
    """
    if not is_available():
        return None
    try:
        import anthropic
        import base64
    except ImportError:
        return None

    client = anthropic.Anthropic()
    model = os.environ.get("LLM_MODEL", "claude-haiku-4-5-20251001")
    schema_str = json.dumps(INVOICE_SCHEMA, indent=2, ensure_ascii=False)
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    prompt = (
        "To zdjęcie faktury. Odczytaj dane i wyodrębnij zgodnie ze schematem.\n\n"
        "ZASADY:\n"
        "- Zwróć WYŁĄCZNIE poprawny JSON, bez komentarzy i bez ```.\n"
        "- Przypisuj NIP i adres do właściwej strony (sprzedawca vs nabywca) "
        "na podstawie układu i nagłówków.\n"
        "- Kwoty jako liczby (kropka dziesiętna). Daty w formacie YYYY-MM-DD.\n"
        "- Brakujące pola ustaw na null. Nie zgaduj.\n\n"
        "ZASADY DLA KWOT ZBIORCZYCH (ważne):\n"
        "- total_amount = kwota brutto do zapłaty ('razem').\n"
        "- net_amount = suma wartości NETTO pozycji.\n"
        "- tax_amount = kwota VAT. Jeśli brak wprost, policz: "
        "total_amount - net_amount. NIE kopiuj tu netto pojedynczej pozycji "
        "ani stawki procentowej.\n"
        "- Kolumna 'VAT' w tabeli to zwykle STAWKA procentowa (np. 23%), "
        "nie kwota podatku.\n"
        "- Sprawdź: net_amount + tax_amount = total_amount.\n\n"
        f"SCHEMAT:\n{schema_str}"
    )

    try:
        resp = client.messages.create(
            model=model,
            max_tokens=1500,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {
                        "type": "base64", "media_type": media_type, "data": b64}},
                    {"type": "text", "text": prompt},
                ],
            }],
        )
        raw = "".join(b.text for b in resp.content if b.type == "text").strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return json.loads(raw)
    except Exception:  # noqa: BLE001
        return None
