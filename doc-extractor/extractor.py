"""
Silnik ekstrakcji danych z dokumentów (faktury, PDF-y, formularze).

Strategia warstwowa:
1. Warstwa tekstowa: pdfplumber wyciąga natywny tekst i tabele z PDF.
2. Fallback OCR: jeśli strona nie ma warstwy tekstowej (skan), renderujemy
   ją do obrazu i puszczamy przez Tesseract.
3. Warstwa parsowania pól: reguły + wyrażenia regularne wyciągają
   ustrukturyzowane pola (NIP, kwoty, daty, numer faktury...).

W realnym wdrożeniu warstwę 3 zastępuje/uzupełnia LLM z wymuszonym
schematem JSON (function calling / structured output) - patrz README.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Any

import pdfplumber
from PIL import Image
import pytesseract


# ---------------------------------------------------------------------------
# Model danych wyjściowych
# ---------------------------------------------------------------------------

@dataclass
class LineItem:
    description: str
    quantity: str | None = None
    unit_price: str | None = None
    total: str | None = None
    vat_rate: str | None = None
    unit: str | None = None


@dataclass
class ExtractionResult:
    source_file: str
    page_count: int
    extraction_method: str          # "text", "ocr", lub "hybrid"
    document_type: str              # "invoice", "form", "generic"
    fields: dict[str, Any] = field(default_factory=dict)
    line_items: list[LineItem] = field(default_factory=list)
    raw_text: str = ""
    confidence: float = 0.0
    warnings: list[str] = field(default_factory=list)
    field_warnings: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        d = asdict(self)
        return d


# ---------------------------------------------------------------------------
# Warstwa 1 + 2: pozyskanie tekstu
# ---------------------------------------------------------------------------

def _ocr_page(page) -> str:
    """Renderuje stronę PDF do obrazu i wykonuje OCR."""
    # 200 DPI zamiast 300 - wyraźnie niższe zużycie pamięci przy porównywalnej
    # jakości rozpoznania; istotne na darmowych planach z limitem 512 MB RAM.
    pil_image = page.to_image(resolution=200).original
    if not isinstance(pil_image, Image.Image):
        pil_image = Image.open(io.BytesIO(pil_image))
    # pol+eng - polskie i angielskie faktury
    langs = "eng"
    try:
        return pytesseract.image_to_string(pil_image, lang="pol+eng")
    except pytesseract.TesseractError:
        return pytesseract.image_to_string(pil_image, lang=langs)


def extract_text(pdf_bytes: bytes) -> tuple[str, int, str, list[str]]:
    """
    Zwraca (pełny_tekst, liczba_stron, metoda, ostrzeżenia).
    Metoda: 'text' | 'ocr' | 'hybrid'
    """
    warnings: list[str] = []
    pages_text: list[str] = []
    methods: set[str] = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_count = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            native = (page.extract_text() or "").strip()
            if len(native) >= 20:          # strona ma sensowną warstwę tekstową
                pages_text.append(native)
                methods.add("text")
            else:
                try:
                    ocr = _ocr_page(page).strip()
                    if ocr:
                        pages_text.append(ocr)
                        methods.add("ocr")
                    else:
                        warnings.append(f"Strona {i+1}: brak tekstu i pusty OCR.")
                except Exception as exc:            # noqa: BLE001
                    warnings.append(f"Strona {i+1}: OCR nieudany ({exc}).")

    if methods == {"text"}:
        method = "text"
    elif methods == {"ocr"}:
        method = "ocr"
    elif methods:
        method = "hybrid"
    else:
        method = "none"

    return "\n\n".join(pages_text), page_count, method, warnings


# ---------------------------------------------------------------------------
# Warstwa 3: parsowanie pól (reguły / regex)
# ---------------------------------------------------------------------------

# Wzorce kwot - obsługa formatów międzynarodowych:
#   2 676,00  (PL: spacja=tysiące, przecinek=dziesiętne)
#   2,676.00  (US/UK: przecinek=tysiące, kropka=dziesiętne)
#   1.234,56  (DE: kropka=tysiące, przecinek=dziesiętne)
#   1190,00 / 1000.00  (bez separatora tysięcy)
#   999  (całkowita)
# Kolejność: najpierw warianty z separatorem tysięcy (najdłuższe), potem prostsze.
_MONEY = (
    r"(\d{1,3}(?:,\d{3})+\.\d{2}"                    # 2,676.00 (US)
    r"|\d{1,3}(?:\.\d{3})+,\d{2}"                     # 1.234,56 (DE)
    r"|\d{1,3}(?:[ \u00a0]\d{3})+(?:[.,]\d{2})?"      # 2 676,00 (PL, spacja)
    r"|\d+[.,]\d{2}"                                  # 1190,00 bez separatora
    r"|\d+)"                                          # 999 całkowita
)
_DATE_PATTERNS = [
    r"\d{4}-\d{2}-\d{2}",
    r"\d{2}[./]\d{2}[./]\d{4}",
    r"\d{2}\.\d{2}\.\d{2}\b",
]


def _find_first(patterns: list[str], text: str, flags=re.I) -> str | None:
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            return m.group(0)
    return None


def _find_group(pattern: str, text: str, flags=re.I, group: int = 1) -> str | None:
    m = re.search(pattern, text, flags)
    if not m:
        return None
    # jeśli wzorzec ma kilka grup alternatywnych (a|b), bierzemy pierwszą niepustą
    if m.lastindex and m.lastindex > 1:
        for g in m.groups():
            if g:
                return g.strip()
        return None
    val = m.group(group)
    return val.strip() if val else None


def _detect_type(text: str) -> str:
    t = text.lower()
    if any(k in t for k in ("faktura", "invoice", "rachunek", "vat",
                            "rechnung", "facture")):
        return "invoice"
    if any(k in t for k in ("formularz", "form", "wniosek", "application",
                            "formular", "formulaire")):
        return "form"
    return "generic"


def _parse_invoice_fields(text: str) -> dict[str, Any]:
    fields: dict[str, Any] = {}

    # Numer faktury
    fields["invoice_number"] = _find_group(
        r"(?:faktura(?:\s+vat)?|invoice|rechnung|facture)\s*"
        r"(?:nr|no|number|nummer|numéro|numero|#)?[:\s]*"
        r"([A-Z0-9][A-Z0-9/\-\.]{2,})", text
    )

    # NIP-y: może być kilka (sprzedawca + nabywca). Zbieramy wszystkie,
    # a potem przypisujemy do stron na podstawie kolejności/kontekstu.
    all_nips = [re.sub(r"[^\d]", "", n) for n in
                re.findall(r"NIP[:\s]*([0-9\-\s]{10,15})", text, re.I)]
    all_nips = [n for n in all_nips if len(n) == 10]
    if all_nips:
        fields["seller_nip"] = all_nips[0]
        if len(all_nips) > 1:
            fields["buyer_nip"] = all_nips[1]

    # VAT ID (ogólny, np. PL...) - PL/EN/DE/FR etykiety
    vat = re.search(
        r"\b(?:VAT[- ]?ID|USt-IdNr|Steuernummer|N°\s*TVA|TVA)[:.\s]*"
        r"([A-Z]{2}[0-9A-Z]{8,12})", text, re.I)
    if vat:
        fields["vat_id"] = vat.group(1)

    # Daty - wystawienia / wykonania / płatności (PL/EN/DE/FR)
    issue = _find_group(
        r"(?:data\s+wystawienia|wystawieni[ae]|issue\s*date|date\s+of\s+issue|"
        r"rechnungsdatum|ausstellungsdatum|date\s+d['e]\s?émission|date\s+facture)[:\s]*"
        r"(\d{4}-\d{2}-\d{2}|\d{2}[./-]\d{2}[./-]\d{4}|\d{2}[./]\d{2}[./]\d{2})", text)
    if issue:
        fields["issue_date"] = _norm_date(issue)
    done = _find_group(
        r"(?:data\s+)?(?:wykonani[ae]|dostawy|sprzeda[żz]y|delivery\s*date|"
        r"leistungsdatum|lieferdatum|date\s+de\s+livraison)[:\s]*"
        r"(\d{4}-\d{2}-\d{2}|\d{2}[./-]\d{2}[./-]\d{4})", text)
    if done:
        fields["service_date"] = _norm_date(done)
    due = _find_group(
        r"(?:termin\s+p[łl]atno[śs]ci|due\s*date|payment\s+due|"
        r"fällig(?:keitsdatum)?|zahlbar\s+bis|date\s+d['e]\s?échéance|échéance)[:\s]*"
        r"(\d{4}-\d{2}-\d{2}|\d{2}[./-]\d{2}[./-]\d{4})", text)
    if due:
        fields["due_date"] = _norm_date(due)

    # Kwoty (PL/EN/DE/FR)
    total = _find_group(
        r"(?:nale[żz]no[śs][cć]i?\s+razem|razem\s+do\s+zap[łl]aty|do\s+zap[łl]aty|"
        r"total\s+due|amount\s+due|grand\s+total|suma\s+brutto|brutto|"
        r"gesamtbetrag|rechnungsbetrag|zu\s+zahlen|montant\s+total|total\s+ttc|"
        r"net\s+à\s+payer)[:\s]*" + _MONEY, text)
    if total:
        fields["total_amount"] = _clean_money(total)
    net = _find_group(
        r"(?:suma\s+netto|razem\s+netto|net(?:\s+amount)?|subtotal|"
        r"nettobetrag|montant\s+ht|total\s+ht)[:\s]*" + _MONEY, text)
    if net:
        fields["net_amount"] = _clean_money(net)

    # Waluta
    cur = _find_first([r"\bPLN\b", r"\bEUR\b", r"\bUSD\b", r"\bGBP\b", r"€", r"\$"], text)
    if cur:
        fields["currency"] = {"€": "EUR", "$": "USD"}.get(cur, cur)

    # Metoda płatności (PL/EN/DE/FR)
    method = _find_group(
        r"(?:metoda|spos[óo]b|forma)\s+p[łl]atno[śs]ci[:\s]*([^\n]{2,20})|"
        r"(?:payment\s+method|zahlungsart|zahlungsweise|mode\s+de\s+paiement)"
        r"[:\s]*([^\n]{2,20})", text)
    if method:
        fields["payment_method"] = method.strip()

    # IBAN - dowolny kraj (2 litery + 2 cyfry + do 30 znaków)
    iban = re.search(r"\b([A-Z]{2}\d{2}(?:[ ]?[A-Z0-9]){10,30})", text)
    if iban:
        fields["bank_account"] = re.sub(r"\s+", " ", iban.group(1)).strip()
    # BIC / SWIFT
    bic = re.search(r"\b(?:BIC|SWIFT)[:\s]*([A-Z]{4}[A-Z0-9]{4,7})", text)
    if bic:
        fields["bank_bic"] = bic.group(1)

    # Numery referencyjne (PO/.../...)
    refs = re.findall(r"\bPO/\d+/\d+/\d+", text)
    if refs:
        fields["reference_numbers"] = ", ".join(dict.fromkeys(refs))

    # Sprzedawca / nabywca - nazwa (PL/EN/DE/FR)
    seller = _find_group(
        r"(?:sprzedawca|seller|verkäufer|lieferant|vendeur|émetteur)"
        r"[:\s]*\n?\s*([^\n]{3,60})", text)
    if seller:
        fields["seller"] = seller.strip()
    buyer = _find_group(
        r"(?:nabywca|buyer|bill\s+to|käufer|kunde|rechnungsempfänger|"
        r"acheteur|client|destinataire)[:\s]*\n?\s*([^\n]{3,60})", text)
    if buyer:
        fields["buyer"] = buyer.strip()

    # E-mail
    email = re.search(r"[\w.\-]+@[\w.\-]+\.\w+", text)
    if email:
        fields["email"] = email.group(0)

    return fields


def _norm_date(raw: str) -> str:
    """Normalizuje różne formaty dat do YYYY-MM-DD (gdy się da)."""
    raw = raw.strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%y", "%d/%m/%y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return raw


def _clean_money(raw: str) -> str:
    """Normalizuje '1 234,56' / '1,234.56' -> '1234.56'."""
    s = raw.strip().replace("\u00a0", " ").replace(" ", "")
    # jeśli przecinek jest separatorem dziesiętnym
    if re.search(r",\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    return s


def _extract_tables(pdf_bytes: bytes) -> list[LineItem]:
    """Próbuje wyciągnąć pozycje z tabel (kolumny opis/ilość/cena/wartość)."""
    items: list[LineItem] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    if not table or len(table) < 2:
                        continue
                    header = [str(c or "").lower() for c in table[0]]
                    # rozpoznaj kolumny
                    def col(*keys):
                        for idx, h in enumerate(header):
                            if any(k in h for k in keys):
                                return idx
                        return None
                    c_desc = col("opis", "nazwa", "description", "item", "produkt")
                    c_qty = col("ilo", "qty", "quantity", "szt")
                    c_price = col("cena", "price", "unit")
                    c_total = col("warto", "total", "amount", "kwota")
                    if c_desc is None:
                        continue
                    for row in table[1:]:
                        if not row or not (row[c_desc] or "").strip():
                            continue
                        items.append(LineItem(
                            description=str(row[c_desc]).strip(),
                            quantity=str(row[c_qty]).strip() if c_qty is not None and row[c_qty] else None,
                            unit_price=str(row[c_price]).strip() if c_price is not None and row[c_price] else None,
                            total=str(row[c_total]).strip() if c_total is not None and row[c_total] else None,
                        ))
    except Exception:  # noqa: BLE001
        pass
    return items


def _line_items_from_text(text: str) -> list[LineItem]:
    """
    Fallback: gdy tabela nie ma linii siatki, pozycje są zwykłymi wierszami
    tekstu zakończonymi liczbami. Wykrywa wiersze typu:
      'Opis produktu   1   1 500,00   1 500,00'
    Cena i wartość MUSZĄ mieć część dziesiętną (grosze) - to odróżnia
    prawdziwe pozycje od przypadkowych ciągów cyfr (np. numeru konta).
    """
    items: list[LineItem] = []
    money_dec = r"(\d{1,3}(?:[ .,\u00a0]\d{3})*[.,]\d{2})"  # zawsze z groszami
    row_re = re.compile(r"^(.+?)\s+(\d+)\s+" + money_dec + r"\s+" + money_dec + r"\s*$")
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        low = line.lower()
        if any(k in low for k in ("netto", "vat", "razem", "suma", "total",
                                  "do zap", "opis", "ilo", "iban", "bic",
                                  "konto", "pln:", "eur:", "swift")):
            continue
        # pomiń linie wyglądające na numer konta (dużo cyfr z rzędu)
        if re.search(r"\d{4}[ ]\d{4}[ ]\d{4}", line):
            continue
        m = row_re.match(line)
        if m:
            items.append(LineItem(
                description=m.group(1).strip(),
                quantity=m.group(2),
                unit_price=_clean_money(m.group(3)),
                total=_clean_money(m.group(4)),
            ))
    return items


def _confidence(fields: dict, method: str, items: list, engine: str = "rules") -> float:
    """Prosty wskaźnik pewności na bazie ilości znalezionych pól."""
    key_fields = ["invoice_number", "total_amount", "issue_date"]
    seller_ok = bool(fields.get("seller") or fields.get("seller_name"))
    hit = sum(1 for k in key_fields if fields.get(k)) + (1 if seller_ok else 0)
    score = hit / (len(key_fields) + 1)
    if items:
        score = min(1.0, score + 0.1)
    if method == "ocr":
        score *= 0.85
    elif method == "hybrid":
        score *= 0.92
    if engine == "llm":
        score = min(1.0, score + 0.05)   # LLM zwykrywa pola pełniej
    return round(score, 2)


# ---------------------------------------------------------------------------
# Publiczne API
# ---------------------------------------------------------------------------

def extract_document(pdf_bytes: bytes, filename: str = "document.pdf") -> ExtractionResult:
    text, page_count, method, warnings = extract_text(pdf_bytes)
    doc_type = _detect_type(text)

    fields: dict[str, Any] = {}
    line_items: list[LineItem] = []
    engine = "rules"

    if doc_type == "invoice":
        # Ścieżka preferowana: LLM (jeśli jest klucz API). Czyta układ faktury
        # jak człowiek - poprawnie przypisuje NIP-y do stron, radzi sobie z
        # dowolnymi układami tabel. Gdy klucza brak lub błąd -> fallback reguły.
        llm_data = _try_llm(text)
        if llm_data:
            fields, line_items = _map_llm_result(llm_data)
            engine = "llm"
        else:
            fields = _parse_invoice_fields(text)
            line_items = _extract_tables(pdf_bytes)
            if not line_items:
                line_items = _line_items_from_text(text)
    else:
        # dla formularzy/ogólnych: wyciągnij pary klucz: wartość
        for m in re.finditer(r"^([A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż /]{3,40})[:\uFF1A]\s*(.+)$",
                             text, re.M):
            key = m.group(1).strip().lower().replace(" ", "_")
            fields.setdefault(key, m.group(2).strip())
        email = re.search(r"[\w.\-]+@[\w.\-]+\.\w+", text)
        if email:
            fields["email"] = email.group(0)

    result = ExtractionResult(
        source_file=filename,
        page_count=page_count,
        extraction_method=method,
        document_type=doc_type,
        fields=fields,
        line_items=line_items,
        raw_text=text,
        confidence=_confidence(fields, method, line_items, engine),
        warnings=warnings,
    )
    # doklejamy informację, który silnik zadziałał (widoczne w UI)
    result.fields["_engine"] = engine
    _attach_validation(result)
    return result


def _attach_validation(result: "ExtractionResult") -> None:
    """Dokłada wyniki walidacji pól (NIP, kwoty, daty) do rezultatu."""
    try:
        from validators import validate_fields
        items_dicts = [{"total": it.total} for it in result.line_items]
        result.field_warnings = validate_fields(result.fields, items_dicts)
    except Exception:  # noqa: BLE001
        result.field_warnings = {}


def _try_llm(text: str):
    """Bezpieczne wywołanie modułu LLM - nie wywala się, gdy go brak."""
    try:
        from llm_extractor import extract_with_llm
        return extract_with_llm(text)
    except Exception:  # noqa: BLE001
        return None


def _map_llm_result(data: dict) -> tuple[dict, list[LineItem]]:
    """Mapuje odpowiedź LLM na nasze struktury (pola + pozycje)."""
    items = []
    for it in (data.get("line_items") or []):
        if not isinstance(it, dict):
            continue
        desc = it.get("description")
        if not desc:
            continue
        items.append(LineItem(
            description=str(desc),
            quantity=_s(it.get("quantity")),
            unit_price=_s(it.get("unit_price")),
            total=_s(it.get("total")),
            vat_rate=_s(it.get("vat_rate")),
            unit=_s(it.get("unit")),
        ))
    # pola bez line_items (te idą osobno)
    fields = {k: v for k, v in data.items()
              if k != "line_items" and v not in (None, "", [])}
    return fields, items


def _s(v):
    """Bezpieczna konwersja liczby/None na string do wyświetlenia."""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def extract_image(image_bytes: bytes, filename: str, media_type: str) -> ExtractionResult:
    """
    Ekstrakcja z obrazu faktury (zdjęcie z aparatu lub plik JPG/PNG).
    Ścieżka preferowana: Claude czyta obraz natywnie (najlepszy odczyt zdjęć).
    Fallback bez klucza API: Tesseract OCR na obrazie + parser regułowy.
    """
    engine = "rules"
    fields: dict[str, Any] = {}
    line_items: list[LineItem] = []
    method = "image"

    # 1) Preferowana ścieżka: Claude vision (jeśli jest klucz API)
    llm_data = _try_llm_image(image_bytes, media_type)
    if llm_data:
        fields, line_items = _map_llm_result(llm_data)
        engine = "llm"
    else:
        # 2) Fallback: OCR obrazu przez Tesseract, potem reguły
        try:
            img = Image.open(io.BytesIO(image_bytes))
            try:
                text = pytesseract.image_to_string(img, lang="pol+eng")
            except pytesseract.TesseractError:
                text = pytesseract.image_to_string(img, lang="eng")
        except Exception:  # noqa: BLE001
            text = ""
        method = "ocr"
        fields = _parse_invoice_fields(text)
        line_items = _line_items_from_text(text)

    result = ExtractionResult(
        source_file=filename,
        page_count=1,
        extraction_method=method,
        document_type="invoice",
        fields=fields,
        line_items=line_items,
        raw_text="(źródło: obraz)",
        confidence=_confidence(fields, method, line_items, engine),
        warnings=[],
    )
    result.fields["_engine"] = engine
    _attach_validation(result)
    return result


def _try_llm_image(image_bytes: bytes, media_type: str):
    """Bezpieczne wywołanie ekstrakcji obrazu przez LLM."""
    try:
        from llm_extractor import extract_from_image
        return extract_from_image(image_bytes, media_type)
    except Exception:  # noqa: BLE001
        return None


if __name__ == "__main__":
    import sys, json
    with open(sys.argv[1], "rb") as f:
        res = extract_document(f.read(), sys.argv[1])
    print(json.dumps(res.to_dict(), indent=2, ensure_ascii=False))
