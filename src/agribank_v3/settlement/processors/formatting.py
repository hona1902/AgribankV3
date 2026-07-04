from __future__ import annotations

from openpyxl.styles import Alignment, Font


def style_agency_header(
    sheet,
    *,
    start_column: int,
    end_column: int,
) -> None:
    """Apply common agency header typography for settlement reports."""
    for row in range(1, 5):
        cell = sheet.cell(row, start_column)
        cell.font = Font(name="Times New Roman", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")


def style_currency_unit(cell) -> None:
    """Format the currency unit note used in settlement report headers."""
    cell.font = Font(
        name="Times New Roman",
        size=10,
        italic=True,
    )
    cell.alignment = Alignment(
        horizontal="right",
        vertical="center",
    )


def apply_page_number_header(
    sheet,
    *,
    text: str = "&P/&N",
    exclude_first_page: bool = True,
) -> None:
    """Add centered page numbering to settlement report headers."""
    if exclude_first_page:
        sheet.HeaderFooter.differentFirst = True
        sheet.firstHeader.left.text = ""
        sheet.firstHeader.center.text = ""
        sheet.firstHeader.right.text = ""
    sheet.oddHeader.center.text = text


def setup_a4_print_layout(
    sheet,
    *,
    print_area: str,
    orientation: str = "landscape",
    title_rows: str | None = None,
    page_number_header: bool = True,
) -> None:
    """Configure a settlement worksheet so users can print immediately."""
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.orientation = orientation
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = print_area
    if title_rows:
        sheet.print_title_rows = title_rows
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.5
    sheet.page_margins.bottom = 0.5
    sheet.page_margins.header = 0.2
    sheet.page_margins.footer = 0.2
    if page_number_header:
        apply_page_number_header(sheet)
