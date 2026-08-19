# Changelog

Historia rozwoju projektu InvoiceFlow — aplikacji do odczytu danych z faktur.

## Etapy rozwoju

### 1. Utworzenie aplikacji do odczytu faktur
Silnik ekstrakcji danych z PDF (pdfplumber + Tesseract OCR + parser regułowy),
backend Flask z API oraz frontend do wgrywania plików i podglądu wyników.

### 2. Wdrożenie w chmurze (Render)
Konfiguracja pod produkcję: Gunicorn, Docker z Tesseractem, port ze zmiennej
środowiskowej, plik `render.yaml`. Aplikacja dostępna pod publicznym adresem —
każdy może wejść przez przeglądarkę.

### 3. Wgranie projektu na GitHub
Inicjalizacja repozytorium, bezpieczne zarządzanie kluczem API (poza repo,
w zmiennych środowiskowych), czysta historia commitów.

### 4. Integracja z LLM (Claude)
Inteligentny odczyt faktur — model czyta układ dokumentu jak człowiek:
poprawnie rozdziela sprzedawcę i nabywcę, radzi sobie z dowolnymi układami.
Automatyczny fallback na parser regułowy, gdy brak klucza API.

### 5. Edytowalny formularz i eksport do Excela
Możliwość poprawiania odczytanych pól przed zapisem, ze wskaźnikiem pewności
przy każdym polu. Eksport do `.xlsx` oraz JSON.

### 6. Wersja mobilna i skanowanie aparatem
Interfejs dostosowany do dotyku oraz możliwość zrobienia zdjęcia faktury
telefonem lub wyboru z galerii (zdjęcia czyta Claude natywnie).

### 7. Funkcje pod portfolio
- Walidacja pól: suma kontrolna NIP, spójność kwot (netto + VAT = brutto), logika dat.
- Podgląd oryginału obok wyodrębnionych danych.
- Przycisk „wypróbuj na przykładzie".

### 8. Wersje językowe (PL / EN / DE / FR)
Pełne tłumaczenie interfejsu z automatycznym wykrywaniem języka z przeglądarki
i przełącznikiem. Parser regułowy rozszerzony o niemieckie i francuskie faktury.

### 9. Poprawki parsera kwot
Obsługa międzynarodowych formatów liczb (`2 676,00` / `2,676.00` / `1.234,56`)
i wyeliminowanie błędnego wykrywania numerów kont jako pozycji faktury.

### 10. Dopracowanie UX/UI
Ograniczenie szerokości górnego panelu na dużych ekranach oraz zamiana
listy rozwijanej wyboru języka na cztery flagi (PL / GB / DE / FR).

### 11. Eksport analityczny (Excel + Power BI)
- Excel: raport wielozakładkowy — Podsumowanie (KPI, rozbicie VAT, wykres kołowy),
  Pozycje (z formułami), Dane faktury. Kwoty liczone formułami Excela.
- Power BI: płaski plik CSV (jeden wiersz = jedna pozycja) gotowy do wczytania
  i budowy dashboardów; po zebraniu wielu faktur wystarczy je scalić.

### 12. Dopracowanie eksportu Excel
- KPI w kolejności: Suma netto → Stawka VAT → Suma VAT → Kwota brutto → Liczba pozycji.
- Naprawiono przepływ stawki VAT z pozycji (pole vat_rate w modelu LineItem) —
  stawka pojawia się teraz w tabeli pozycji i w rozbiciu VAT.
- Eksport Excel w języku interfejsu (PL/EN/DE/FR): nazwy zakładek, KPI, etykiety.

### 13. Przetwarzanie wielu faktur naraz (batch)
- Wybór wielu plików PDF/obrazów jednocześnie; przetwarzanie sekwencyjne
  z paskiem postępu ("Przetwarzam 3/10").
- Panel zbiorczy: lista zebranych faktur z sumami netto/VAT/brutto,
  możliwość usuwania pojedynczych pozycji.
- Eksport zbiorczy Excel: zakładki Podsumowanie (globalne KPI + wykres wg
  sprzedawcy), Faktury (lista), Pozycje (wszystkie pozycje).
- Eksport zbiorczy CSV pod Power BI: wszystkie pozycje ze wszystkich faktur
  w jednym pliku. Dane trzymane w pamięci przeglądarki (bez zapisu na serwerze).

### 14. Przyspieszenie przetwarzania wielu faktur (równoległość)
- Faktury przetwarzane równolegle (do 4 naraz) zamiast jedna po drugiej.
  Realne przyspieszenie ok. 4x (100 faktur: ~8 min -> ~2 min).
- Panel zbiorczy wypełnia się na żywo, w miarę spływania wyników.
- Backend: zwiększona liczba wątków Gunicorn (2 -> 4), dostrojona pod
  limit pamięci darmowego planu (ekstrakcja LLM to głównie czekanie na sieć).

### 15. Nawigacja między wgranymi fakturami
- Kliknięcie faktury na liście zbiorczej otwiera jej pełne szczegóły
  (dane + podgląd oryginału) w sekcji poniżej.
- Aktywna faktura jest podświetlona na liście.
- Podgląd oryginału zachowywany dla każdej faktury (nie tylko ostatniej).
