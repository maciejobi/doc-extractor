FROM python:3.12-slim

# Tesseract OCR + języki (polski i angielski) potrzebne dla skanów
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr tesseract-ocr-pol tesseract-ocr-eng \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000
# Gunicorn w produkcji. Konfiguracja dostrojona pod darmowy plan (512 MB RAM):
#   -w 1            : 1 worker (2+ workery + OCR potrafią przekroczyć pamięć)
#   --threads 4     : 4 wątki -> obsługa równoległych żądań. Ekstrakcja przez
#                     LLM to głównie CZEKANIE na odpowiedź API (nie obciążenie
#                     RAM/CPU), więc wątki dają realne przyspieszenie batcha
#                     bez wysycania pamięci. Frontend wysyła max 4 naraz.
#   --timeout 180   : długi timeout (batch + OCR skanów)
#   --max-requests  : recykling workera po N żądaniach (czyści wycieki pamięci)
# Port ze zmiennej PORT (chmury przydzielają dynamicznie), fallback 5000 lokalnie.
CMD gunicorn -b 0.0.0.0:${PORT:-5000} \
    -w 1 --threads 4 --timeout 180 \
    --max-requests 80 --max-requests-jitter 15 \
    --access-logfile - --error-logfile - \
    app:app
