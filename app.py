"""Flask app: upload an M-PESA statement PDF, download the transactions as Excel."""

from __future__ import annotations

import io
import os
import uuid

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from excel_writer import build_workbook
from mpesa_pdf import ExtractionError, MpesaPdfError, PasswordError, extract_statement

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", uuid.uuid4().hex)

ALLOWED_EXTENSIONS = {".pdf"}
MAX_CONTENT_MB = 20
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_MB * 1024 * 1024


def _allowed(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", max_mb=MAX_CONTENT_MB)


@app.route("/convert", methods=["POST"])
def convert():
    upload = request.files.get("statement")
    password = (request.form.get("password") or "").strip()

    if upload is None or not upload.filename:
        flash("Please choose a PDF statement to upload.", "error")
        return redirect(url_for("index"))
    if not _allowed(upload.filename):
        flash("Only PDF files are supported.", "error")
        return redirect(url_for("index"))

    data = upload.read()
    if not data:
        flash("The uploaded file is empty.", "error")
        return redirect(url_for("index"))

    try:
        statement = extract_statement(data, password)
        workbook = build_workbook(statement.transactions, secure_filename(upload.filename))
    except PasswordError as exc:
        flash(str(exc), "error")
        return redirect(url_for("index"))
    except (ExtractionError, MpesaPdfError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("index"))
    except Exception:  # noqa: BLE001 - never leak a stack trace to the browser
        app.logger.exception("Unexpected failure while converting %s", upload.filename)
        flash("Something went wrong while processing the file. Please try again.", "error")
        return redirect(url_for("index"))

    for warning in statement.warnings:
        app.logger.info("Statement note: %s", warning)

    base = os.path.splitext(secure_filename(upload.filename))[0] or "mpesa_statement"
    output_name = f"{base}_transactions.xlsx"
    stream = io.BytesIO(workbook)
    stream.seek(0)
    return send_file(
        stream,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=output_name,
    )


@app.errorhandler(413)
def too_large(_error):
    flash(f"File is too large. The limit is {MAX_CONTENT_MB} MB.", "error")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
