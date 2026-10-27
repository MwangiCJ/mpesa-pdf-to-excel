"""Build statements that span pages, with and without a repeated header."""

from __future__ import annotations

import pymupdf as fitz

WIDTHS = [80, 118, 275, 82, 72, 78, 75]
LABELS = [
    "Receipt No.", "Completion Time", "Details", "Transaction Status",
    "Paid In", "Withdrawn", "Balance",
]

PAGE1 = [
    ["RFK10AAAAA", "01/09/2025 08:00:00", "PAYMENT RECEIVED FROM ALICE", "Completed", "1000.00", "", "1000.00"],
    ["RFK10BBBBB", "02/09/2025 09:15:00", "PAYMENT TO 254700111222", "Completed", "", "250.00", "750.00"],
    ["RFK10CCCCC", "03/09/2025 10:30:00", "ATM WITHDRAWAL WESTLANDS", "Completed", "", "500.00", "250.00"],
]
PAGE2 = [
    ["RFK11DDDDD", "04/09/2025 11:45:00", "PAYMENT RECEIVED FROM BOB", "Completed", "3000.00", "", "3250.00"],
    ["RFK11EEEEE", "05/09/2025 12:00:00", "LIPA NA M-PESA 400300", "Completed", "", "1200.00", "2050.00"],
    ["RFK11FFFFF", "06/09/2025 14:20:00", "PAYMENT TO 254700333444", "Completed", "", "300.00", "1750.00"],
]


def _draw_row(page: fitz.Page, y: float, values: list[str], x0: float) -> None:
    x = x0
    for width, value in zip(WIDTHS, values):
        rect = fitz.Rect(x, y, x + width, y + 26)
        page.draw_rect(rect, color=(0, 0, 0), width=0.6)
        if value:
            align = fitz.TEXT_ALIGN_RIGHT if value[0].isdigit() and "/" not in value[:3] else fitz.TEXT_ALIGN_LEFT
            if page.insert_textbox(rect + (4, 6, -4, -4), value, fontsize=8, align=align) < 0:
                raise RuntimeError(f"text did not fit cell: {value!r}")
        x += width


def build_multipage_pdf(with_repeated_header: bool) -> bytes:
    """Two pages of transactions; page 2 optionally repeats the header row."""
    doc = fitz.open()
    x0 = 40.0
    assert sum(WIDTHS) <= 842 - x0 - 10

    for page_index, rows in enumerate((PAGE1, PAGE2)):
        page = doc.new_page(width=842, height=595)
        page.insert_text((40, 45), "M-PESA STATEMENT - PAGE %d" % (page_index + 1), fontsize=14)
        y = 70.0
        if page_index == 0 or with_repeated_header:
            _draw_row(page, y, LABELS, x0)
            y += 26
        for row in rows:
            _draw_row(page, y, row, x0)
            y += 26

    payload = doc.tobytes(
        encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="pw123", owner_pw="pw123"
    )
    doc.close()
    return payload


if __name__ == "__main__":
    from mpesa_pdf import extract_statement

    for flag in (True, False):
        payload = build_multipage_pdf(with_repeated_header=flag)
        statement = extract_statement(payload, "pw123")
        frame = statement.transactions
        print(f"repeated_header={flag}: rows={len(frame)} pages={statement.pages} "
              f"cols={len(frame.columns)}")
        print(frame[["receipt_no", "paid_in", "withdrawn", "balance"]].to_string(index=False))
        expected = len(PAGE1) + len(PAGE2)
        assert len(frame) == expected, f"expected {expected} rows"
    print("OK")
