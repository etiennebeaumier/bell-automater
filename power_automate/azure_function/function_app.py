"""Azure Function: BCECN PDF parser endpoint.

Triggered by HTTP POST from Power Automate.
Accepts a base64-encoded PDF, detects the bank, runs the appropriate parser,
and returns structured JSON for all yield/spread fields.

Deploy this function alongside the parsers/ package from the root of the repo.
The directory layout expected on Azure is:

    function_app.py
    parsers/
        __init__.py
        td.py
        scotiabank.py
        cibc.py
        nbcm.py
        bmo.py
    host.json
    requirements.txt

Power Automate calls this function via the HTTP action:
    POST https://<app>.azurewebsites.net/api/parse?code=<function-key>
    Content-Type: application/json
    Body: { "pdf_base64": "...", "filename": "BCECN 03.02.26 TD.pdf" }
"""

import azure.functions as func
import base64
import json
import logging
import os
import tempfile
from datetime import datetime

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


# ---------------------------------------------------------------------------
# Bank detection (mirrors main.py logic, operates on already-opened PDF text)
# ---------------------------------------------------------------------------

def _detect_bank(filename_lower: str, first_page_text: str) -> str:
    """Return a bank key from filename hint first, then PDF text content."""
    if "scotiabank" in filename_lower or "bns" in filename_lower:
        return "scotiabank"
    if "cibc" in filename_lower:
        return "cibc"
    if "nbcm" in filename_lower or "national bank" in filename_lower:
        return "nbcm"
    if "bmo" in filename_lower:
        return "bmo"
    if "td" in filename_lower:
        return "td"

    t = first_page_text.lower()
    if "scotiabank" in t or "bns-internal" in t:
        return "scotiabank"
    if "cibc capital markets" in t or "cibc" in t:
        return "cibc"
    if "national bank" in t or "nbcm" in t:
        return "nbcm"
    if "bmo nesbitt burns" in t or "bmo capital markets" in t:
        return "bmo"
    if "td securities" in t:
        return "td"

    raise ValueError("Could not detect bank from filename or PDF content.")


# ---------------------------------------------------------------------------
# Parser registry
# ---------------------------------------------------------------------------

def _get_parser(bank_key: str):
    if bank_key == "td":
        from parsers.td import parse_td_pdf
        return parse_td_pdf
    if bank_key == "scotiabank":
        from parsers.scotiabank import parse_scotiabank_pdf
        return parse_scotiabank_pdf
    if bank_key == "cibc":
        from parsers.cibc import parse_cibc_pdf
        return parse_cibc_pdf
    if bank_key == "nbcm":
        from parsers.nbcm import parse_nbcm_pdf
        return parse_nbcm_pdf
    if bank_key == "bmo":
        from parsers.bmo import parse_bmo_pdf
        return parse_bmo_pdf
    raise ValueError(f"No parser registered for bank key: {bank_key!r}")


# ---------------------------------------------------------------------------
# HTTP endpoint
# ---------------------------------------------------------------------------

@app.route(route="parse", methods=["POST"])
def parse_bcecn_pdf(req: func.HttpRequest) -> func.HttpResponse:
    """Parse a BCECN PDF and return structured pricing data as JSON.

    Request body (JSON):
        pdf_base64  str   Required. Base64-encoded PDF file bytes.
        filename    str   Optional. Original filename; used as the primary
                          bank-detection signal (e.g. "BCECN 03.02.26 TD.pdf").

    Response body (JSON) on success (HTTP 200):
        date            str   ISO-8601 date string, e.g. "2026-03-02"
        bank            str   Detected bank name, e.g. "TD"
        cad_spread_3y   float Spread in bps
        cad_yield_3y    float Re-offer yield as decimal (e.g. 0.0450 = 4.50%)
        ...             (all 28 metric fields from COLUMN_MAP in excel_writer.py)

    Response body (JSON) on error (HTTP 400 / 500):
        error           str   Human-readable error message
    """
    # --- Parse request body --------------------------------------------------
    try:
        body = req.get_json()
    except Exception:
        return _error_response("Request body must be valid JSON.", 400)

    pdf_b64: str = body.get("pdf_base64", "")
    filename: str = body.get("filename", "attachment.pdf")

    if not pdf_b64:
        return _error_response("Missing required field: pdf_base64.", 400)

    # --- Decode PDF bytes ----------------------------------------------------
    try:
        pdf_bytes = base64.b64decode(pdf_b64)
    except Exception as exc:
        return _error_response(f"pdf_base64 is not valid base64: {exc}", 400)

    # Write to a temp file so pdfplumber can open it by path (same as local tool)
    tmp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        # --- Detect bank -----------------------------------------------------
        import pdfplumber

        with pdfplumber.open(tmp_path) as pdf:
            if not pdf.pages:
                return _error_response("PDF has no pages.", 422)
            first_page_text = pdf.pages[0].extract_text() or ""

        bank_key = _detect_bank(filename, first_page_text)
        parser = _get_parser(bank_key)

        # --- Parse -----------------------------------------------------------
        data: dict = parser(tmp_path)

        # --- Serialize -------------------------------------------------------
        # datetime → ISO-8601 string; None stays null (JSON serializable).
        output = {
            k: (v.strftime("%Y-%m-%d") if isinstance(v, datetime) else v)
            for k, v in data.items()
        }

        return func.HttpResponse(
            body=json.dumps(output),
            status_code=200,
            mimetype="application/json",
        )

    except ValueError as exc:
        logging.warning("Parse error (client-side): %s", exc)
        return _error_response(str(exc), 422)

    except Exception as exc:
        logging.exception("Unexpected error parsing PDF")
        return _error_response(f"Internal error: {exc}", 500)

    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _error_response(message: str, status_code: int) -> func.HttpResponse:
    return func.HttpResponse(
        body=json.dumps({"error": message}),
        status_code=status_code,
        mimetype="application/json",
    )
