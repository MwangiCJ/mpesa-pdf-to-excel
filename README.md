# M-PESA Statement to Excel

A small Flask web app that turns a password-protected Safaricom **M-PESA statement PDF**
into a formatted **Excel** file.

1. Upload the statement PDF.
2. Enter the PDF password (Safaricom sends it by SMS; it is often your ID number).
3. Download `<statement>_transactions.xlsx`.

## Run it: > `.venv\Scripts\python.exe app.py → http://127.0.0.1:5000`
How it works
1. Upload + password — PyMuPDF opens the PDF from memory; if needs_pass, the typed password is authenticated. Wrong/missing password gives a clear flash message, never a stack trace.
2. Parsing — detects the statement table, maps header cells to canonical fields (receipt_no, completion_time, details, status, paid_in, withdrawn, balance, plus transaction_cost/balance_after when present), and keeps only rows with a parseable date. Columns the statement lacks are dropped.
3. Excel — frozen header, autofilter, alternating rows, dd/mm/yyyy hh:mm:ss timestamps, #,##0.00 money, plus a Summary sheet with totals, period and net cash flow.

## Features

- Unlocks AES-encrypted statement PDFs with the password you type.
- Detects the statement table and reads the standard columns:
  `Receipt No.`, `Completion Time`, `Details`, `Transaction Status`,
  `Paid In`, `Withdrawn`, `Balance`, plus `Transaction Cost` / `Balance After`
  when the statement has them. Columns the statement lacks are dropped.
- Falls back to x/y text coordinates when the PDF has no ruled table lines.
- Produces a styled workbook: frozen header, autofilter, alternating rows,
  `dd/mm/yyyy hh:mm:ss` timestamps and `#,##0.00` money formatting.
- Adds a **Summary** sheet: source file, period, transaction count,
  total received, total withdrawn, total costs and net cash flow.
- Files are read from memory and never written to disk.

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows:  source .venv/bin/activate on macOS/Linux
pip install -r requirements.txt
```

## Run

```bash
python app.py
```

Open <http://127.0.0.1:5000>.

For a production run use a WSGI server, e.g. `pip install gunicorn` then
`gunicorn -w 2 'app:app'` (on Windows use `waitress-serve app:app`).

## Tests

```bash
pip install -r requirements-dev.txt   # adds `requests` for the HTTP checks
python test_pipeline.py               # parsing, Excel output, Flask test client
python test_multipage.py              # multi-page statements, repeated header or not
python app.py &                       # then, in another shell:
python test_http.py                   # live checks against http://127.0.0.1:5000
```

`test_pipeline.py` builds a sample password-protected statement plus an unruled
one, then checks password handling, column parsing, Excel output and the HTTP
endpoints via Flask's test client.

## Project layout

| File              | Purpose                                                  |
| ----------------- | -------------------------------------------------------- |
| `app.py`          | Flask routes: `/` upload page, `/convert` returns the xlsx |
| `mpesa_pdf.py`    | Opens/unlocks the PDF and extracts transaction rows       |
| `excel_writer.py` | Builds the styled workbook                                |
| `templates/index.html` | Upload page (drag & drop, password field)            |
| `test_pipeline.py`| End-to-end checks                                        |

## Configuration

| Variable   | Effect                                              |
| ---------- | --------------------------------------------------- |
| `SECRET_KEY` | Flask session secret (random per start if unset)  |
| `MAX_CONTENT_LENGTH` | Upload limit, default **20 MB**              |

## Troubleshooting

- **"Incorrect password"** — use the password Safaricom sent for *this* statement;
  each statement gets a new one. It is commonly your ID number or your date of birth.
- **"No transactions were found"** — make sure the file is an M-PESA statement PDF.
  Scanned/image-only PDFs contain no extractable text.
- **Nothing downloads** — check the browser allows downloads from `127.0.0.1:5000`.
