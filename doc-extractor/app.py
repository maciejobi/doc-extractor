"""
Serwer aplikacji do ekstrakcji danych z dokumentów.

Endpointy:
  GET  /              -> interfejs (upload pliku)
  POST /api/extract   -> przyjmuje PDF (multipart), zwraca JSON z ekstrakcji
  GET  /api/health    -> status
"""

from __future__ import annotations

import traceback

from flask import Flask, request, jsonify, send_from_directory, Response

from extractor import extract_document, extract_image
from excel_export import build_xlsx
from powerbi_export import build_powerbi_csv
from batch_export import build_batch_xlsx, build_batch_csv

app = Flask(__name__, static_folder="static", static_url_path="")

MAX_MB = 15
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024

# rozpoznawanie typu pliku po rozszerzeniu
IMAGE_TYPES = {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".png": "image/png", ".webp": "image/webp",
    ".heic": "image/jpeg", ".heif": "image/jpeg",  # iPhone (traktujemy jak jpeg)
}


@app.get("/")
def index():
    return send_from_directory("static", "index.html")


@app.get("/api/health")
def health():
    return jsonify(status="ok", service="doc-extractor", version="1.0.0")


@app.post("/api/extract")
def extract():
    if "file" not in request.files:
        return jsonify(error="Brak pliku. Wyślij pole 'file'."), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify(error="Pusta nazwa pliku."), 400

    name = f.filename.lower()
    ext = "." + name.rsplit(".", 1)[-1] if "." in name else ""

    try:
        data = f.read()
        if not data:
            return jsonify(error="Plik jest pusty."), 400

        if ext == ".pdf":
            result = extract_document(data, f.filename)
        elif ext in IMAGE_TYPES:
            result = extract_image(data, f.filename, IMAGE_TYPES[ext])
        else:
            return jsonify(
                error="Obsługiwane formaty: PDF oraz zdjęcia (JPG, PNG, WEBP, HEIC)."
            ), 415

        return jsonify(result.to_dict())
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        return jsonify(error=f"Ekstrakcja nieudana: {exc}"), 500


@app.get("/api/sample")
def sample():
    """Zwraca wynik ekstrakcji dla przykładowej faktury (przycisk 'wypróbuj')."""
    import os
    path = os.path.join("samples", "faktura_cargo.pdf")
    if not os.path.exists(path):
        return jsonify(error="Brak przykładowej faktury."), 404
    try:
        with open(path, "rb") as f:
            result = extract_document(f.read(), "przyklad_faktura.pdf")
        return jsonify(result.to_dict())
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        return jsonify(error=f"Nie udało się wczytać przykładu: {exc}"), 500


@app.get("/samples/<path:fname>")
def sample_file(fname):
    """Serwuje przykładowy plik (do podglądu w przycisku 'wypróbuj')."""
    return send_from_directory("samples", fname)


@app.post("/api/export/xlsx")
def export_xlsx():
    """Przyjmuje dane faktury (JSON, po edycji użytkownika) i zwraca plik Excel."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify(error="Brak danych do eksportu."), 400
    lang = (data.get("lang") or "pl").lower()[:2]
    try:
        xlsx_bytes = build_xlsx(data, lang)
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        return jsonify(error=f"Eksport nieudany: {exc}"), 500

    name = (data.get("source_file") or "faktura").rsplit(".", 1)[0]
    return Response(
        xlsx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{name}.xlsx"'},
    )


@app.post("/api/export/powerbi")
def export_powerbi():
    """Przyjmuje dane faktury i zwraca płaski CSV pod Power BI (1 wiersz = 1 pozycja)."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify(error="Brak danych do eksportu."), 400
    try:
        csv_bytes = build_powerbi_csv(data)
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        return jsonify(error=f"Eksport nieudany: {exc}"), 500
    name = (data.get("source_file") or "faktura").rsplit(".", 1)[0]
    return Response(
        csv_bytes,
        mimetype="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{name}_powerbi.csv"'},
    )


@app.post("/api/export/batch/xlsx")
def export_batch_xlsx():
    """Zbiorczy Excel z wielu faktur."""
    data = request.get_json(silent=True)
    invoices = (data or {}).get("invoices")
    if not invoices:
        return jsonify(error="Brak faktur do eksportu."), 400
    lang = (data.get("lang") or "pl").lower()[:2]
    try:
        xlsx_bytes = build_batch_xlsx(invoices, lang)
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        return jsonify(error=f"Eksport nieudany: {exc}"), 500
    return Response(
        xlsx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="zestawienie_faktur.xlsx"'},
    )


@app.post("/api/export/batch/powerbi")
def export_batch_powerbi():
    """Zbiorczy CSV pod Power BI z wielu faktur."""
    data = request.get_json(silent=True)
    invoices = (data or {}).get("invoices")
    if not invoices:
        return jsonify(error="Brak faktur do eksportu."), 400
    try:
        csv_bytes = build_batch_csv(invoices)
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        return jsonify(error=f"Eksport nieudany: {exc}"), 500
    return Response(
        csv_bytes,
        mimetype="text/csv",
        headers={"Content-Disposition": 'attachment; filename="zestawienie_faktur_powerbi.csv"'},
    )


@app.errorhandler(413)
def too_large(_):
    return jsonify(error=f"Plik przekracza limit {MAX_MB} MB."), 413


if __name__ == "__main__":
    import os
    # Chmury przydzielają port przez zmienną PORT; lokalnie fallback 5000.
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
