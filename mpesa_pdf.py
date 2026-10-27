"""Read a password-protected M-PESA statement PDF and extract transactions."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pymupdf as fitz  # PyMuPDF
import pandas as pd

# Canonical field names for the columns found in Safaricom M-PESA statements.
HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "receipt_no": ("receipt no", "receipt"),
    "completion_time": ("completion time", "date/time", "date"),
    "details": ("details", "description"),
    "status": ("transaction status", "status"),
    "paid_in": ("paid in",),
    "withdrawn": ("withdrawn", "withdraw"),
    "balance": ("balance",),
    "transaction_cost": ("transaction cost", "cost"),
    "balance_after": ("balance after",),
}

# Output column order (only columns present in the statement are kept).
COLUMNS: list[str] = list(HEADER_ALIASES.keys())

DATE_FORMATS = (
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%d/%m/%Y",
    "%d-%m-%Y",
)

AMOUNT_RE = re.compile(r"(?i)\b(kshs?|kes|k)\.?")


class MpesaPdfError(Exception):
    """Base class for user-facing statement reading errors."""


class PasswordError(MpesaPdfError):
    """The PDF is encrypted and the supplied password is missing/wrong."""


class ExtractionError(MpesaPdfError):
    """The PDF was readable but no transaction table could be found."""


@dataclass
class Statement:
    """Result of parsing one M-PESA statement PDF."""

    transactions: pd.DataFrame
    pages: int = 0
    warnings: list[str] = field(default_factory=list)


def _normalise_header(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip().lower().rstrip(":").rstrip(".")


def map_header(cells: list[Optional[str]]) -> dict[int, str]:
    """Map column index -> canonical field name for a header row."""
    mapping: dict[int, str] = {}
    for idx, cell in enumerate(cells):
        normalised = _normalise_header(cell or "")
        if not normalised:
            continue
        for field_name, aliases in HEADER_ALIASES.items():
            if normalised in aliases and field_name not in mapping.values():
                mapping[idx] = field_name
                break
    return mapping


def looks_like_header(mapping: dict[str, ...]) -> bool:
    """A usable M-PESA header must identify the core columns."""
    names = set(mapping.values()) if isinstance(mapping, dict) else set()
    return {"completion_time", "details"}.issubset(names) and len(names) >= 4


def parse_amount(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    text = AMOUNT_RE.sub("", str(value)).replace(",", "").strip()
    if not text or text in {"-", "--", "n/a", "n.a."}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _looks_like_date(text: str) -> bool:
    """Guard against fuzzy parsers inventing a date from a bare time or word."""
    if re.search(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", text):
        return True
    return bool(re.search(r"(?i)\b(\d{1,2}\s+)?(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{2,4}", text))


def parse_datetime(value: Optional[str]) -> Optional[pd.Timestamp]:
    if value is None:
        return None
    text = str(value).strip()
    if not text or not _looks_like_date(text):
        return None
    for fmt in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else pd.Timestamp(parsed)


def open_statement(data: bytes, password: str = "") -> fitz.Document:
    """Open PDF bytes, unlocking them with the user's password if required."""
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:  # noqa: BLE001 - surface as a friendly message
        raise ExtractionError("The file could not be read as a PDF.") from exc

    if doc.needs_pass:
        if not password:
            doc.close()
            raise PasswordError("This PDF is password protected. Please enter the password.")
        if doc.authenticate(password) <= 0:
            doc.close()
            raise PasswordError("Incorrect password. Please check the password and try again.")
    return doc


def _is_data_row(record: dict[str, Optional[str]]) -> bool:
    """Keep only rows that carry a parseable completion time."""
    return parse_datetime(record.get("completion_time")) is not None


def _clean_record(record: dict[str, Optional[str]]) -> dict[str, object]:
    clean: dict[str, object] = {name: record.get(name) for name in COLUMNS}
    clean["completion_time"] = parse_datetime(record.get("completion_time"))
    for amount_col in ("paid_in", "withdrawn", "balance", "transaction_cost", "balance_after"):
        clean[amount_col] = parse_amount(record.get(amount_col))
    for text_col in ("receipt_no", "details", "status"):
        value = clean.get(text_col)
        clean[text_col] = re.sub(r"\s+", " ", str(value)).strip() if value else ""
    return clean


def _cell_text(page: fitz.Page, cell: object) -> Optional[str]:
    """Read a table cell via a clip rect (keeps spaces, unlike Table.extract)."""
    if cell is None:
        return None
    rect = fitz.Rect(cell)
    if rect.is_empty or rect.is_infinite:
        return None
    return page.get_text("text", clip=rect).strip() or None


def _header_columns(page: fitz.Page, row: object) -> list[tuple[float, str]]:
    """Column boundaries from a header row's words.

    Uses word x-positions so a header cell that merged two labels
    (e.g. "Transaction Status Paid In") still yields both columns.
    """
    words = [(w[0], w[4]) for w in page.get_text("words", clip=fitz.Rect(row.bbox))]  # type: ignore[attr-defined]
    return _header_positions(sorted(words))


def _record_from_cells(
    page: fitz.Page, row: object, columns: list[tuple[float, str]]
) -> dict[str, Optional[str]]:
    """Build a record by assigning each cell to the column its left edge falls in."""
    starts = [x for x, _name in columns]
    buckets: dict[str, list[str]] = {name: [] for _x, name in columns}
    for cell in row.cells:  # type: ignore[attr-defined]
        if cell is None:
            continue
        rect = fitz.Rect(cell)
        text = page.get_text("text", clip=rect).strip()
        if not text:
            continue
        slot = max((i for i, start in enumerate(starts) if rect.x0 >= start - 6), default=0)
        buckets[columns[slot][1]].append(text)
    return {name: " ".join(parts) for name, parts in buckets.items() if parts}


def _extract_with_tables(doc: fitz.Document, warnings: list[str]) -> list[dict[str, object]]:
    """Primary strategy: detect ruled tables page by page."""
    records: list[dict[str, object]] = []
    mapping: dict[int, str] = {}

    for page in doc:
        for table in page.find_tables().tables:
            rows = list(table.rows)
            if not rows:
                continue
            grids = [[_cell_text(page, cell) for cell in row.cells] for row in rows]

            header_index = -1
            for idx, grid in enumerate(grids[:3]):  # headers repeat atop each page
                candidate = map_header(grid)
                if looks_like_header(candidate):
                    mapping, header_index = candidate, idx
                    break

            if header_index < 0:
                if not mapping:
                    continue  # no column layout known yet for this document
                widest = max((len(g) for g in grids if g), default=0)
                if widest <= max(mapping):
                    continue  # unrelated table, not a statement continuation
                body = list(zip(rows, grids))
                use_x = False
            else:
                body = list(zip(rows[header_index + 1 :], grids[header_index + 1 :]))
                # Prefer x-based columns when they recover more labels than the
                # cell grid did (a merged header cell hides columns from index mapping).
                x_columns = _header_columns(page, rows[header_index])
                use_x = len({name for _x, name in x_columns}) > len(mapping) and looks_like_header(
                    {name: name for _x, name in x_columns}
                )

            for row, grid in body:
                if use_x:
                    record = _record_from_cells(page, row, x_columns)
                else:
                    record = {
                        name: (grid[i] if i < len(grid) else None)
                        for i, name in mapping.items()
                    }
                if _is_data_row(record):
                    records.append(_clean_record(record))

    if not records:
        warnings.append("No ruled table found; fell back to text extraction.")
    return records


def _header_positions(words: list[tuple[float, str]]) -> list[tuple[float, str]]:
    """Find (x_position, field_name) pairs by matching aliases against word runs."""
    positions: list[tuple[float, str]] = []
    idx = 0
    while idx < len(words):
        matched = False
        for size in (4, 3, 2, 1):
            phrase = " ".join(word for _x, word in words[idx : idx + size]).lower().rstrip(":")
            field_name = next(
                (name for name, aliases in HEADER_ALIASES.items() if phrase in aliases), None
            )
            if field_name and all(field_name != known for _x, known in positions):
                positions.append((words[idx][0], field_name))
                idx += size
                matched = True
                break
        if not matched:
            idx += 1
    return sorted(positions)


def _line_record(
    columns: list[tuple[float, str]], parts: list[tuple[float, str]]
) -> Optional[dict[str, Optional[str]]]:
    """Assign each word of a data line to the column its x-coordinate falls into."""
    bounds = [x for x, _name in columns]
    buckets: dict[str, list[str]] = {name: [] for _x, name in columns}
    for x0, word in parts:
        slot = max((i for i, start in enumerate(bounds) if x0 >= start), default=0)
        buckets[columns[slot][1]].append(word)
    record = {name: " ".join(words) for name, words in buckets.items() if words}
    return record or None


def _visual_lines(page: fitz.Page, y_tolerance: float = 4.0) -> list[list[tuple[float, str]]]:
    """Group page words into visual rows, then sort each row left to right.

    Clustering on y (rather than PyMuPDF's block/line ids) keeps cells written
    by separate drawing calls on the same baseline together.
    """
    words = page.get_text("words")  # x0, y0, x1, y1, word, block, line, word_no
    if not words:
        return []
    ordered = sorted(words, key=lambda w: (w[1], w[0]))

    rows: list[list[tuple[float, str]]] = []
    row_y: list[float] = []  # most recent y per row; drift stays tiny within a line
    for x0, y0, _x1, _y1, word, *_rest in ordered:
        if rows and abs(y0 - row_y[-1]) <= y_tolerance:
            rows[-1].append((x0, word))
            row_y[-1] = y0
        else:
            rows.append([(x0, word)])
            row_y.append(y0)

    return [sorted(row, key=lambda w: w[0]) for row in rows]


def _extract_with_text(doc: fitz.Document, warnings: list[str]) -> list[dict[str, object]]:
    """Fallback: derive column boundaries from text coordinates."""
    records: list[dict[str, object]] = []
    columns: list[tuple[float, str]] = []

    for page in doc:
        for parts in _visual_lines(page):
            text = " ".join(word for _x, word in parts)
            found = _header_positions(parts)
            if looks_like_header({name: name for _x, name in found}):
                columns = found
                continue
            if not columns:
                continue
            record = _line_record(columns, parts)
            if record and _is_data_row(record):
                records.append(_clean_record(record))
            elif "M-PESA" in text.upper():
                warnings.append(text.strip())

    if not records:
        warnings.append("Text extraction produced no rows.")
    return records


def _has_values(series: pd.Series) -> bool:
    """True when a column holds at least one usable value."""
    remaining = series.dropna()
    if remaining.empty:
        return False
    if not pd.api.types.is_object_dtype(remaining) and not pd.api.types.is_string_dtype(remaining):
        return True  # numeric / datetime columns: any surviving value counts
    return remaining.astype(str).str.strip().ne("").any()


def extract_statement(data: bytes, password: str = "") -> Statement:
    """Parse an M-PESA statement PDF into a DataFrame of transactions."""
    doc = open_statement(data, password)
    warnings: list[str] = []
    try:
        pages = doc.page_count
        records = _extract_with_tables(doc, warnings)
        if not records:
            warnings.clear()
            records = _extract_with_text(doc, warnings)
    finally:
        doc.close()

    if not records:
        raise ExtractionError(
            "No transactions were found in the PDF. Please upload an M-PESA statement PDF."
        )

    frame = pd.DataFrame.from_records(records, columns=COLUMNS)
    frame = frame[frame["completion_time"].notna()]
    frame = frame.sort_values("completion_time", kind="stable").reset_index(drop=True)
    present = [col for col in COLUMNS if _has_values(frame[col])]
    frame = frame[present]
    return Statement(transactions=frame, pages=pages, warnings=warnings)
