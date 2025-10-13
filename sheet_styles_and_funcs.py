from openpyxl.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    Font,
    Side,
    Border,
    Alignment
)


dims = {
    "A" : 30,
    "B" : 10,
    "C" : 20,
    "D" : 15,
    "E" : 45,
}

font_style = Font(
    name="Times New Roman",
    size=11,
    bold=False,
)

font_style_bold = Font(
    name="Times New Roman",
    size=11,
    bold=True,
)

thick = Side(
    border_style="medium",
    color="FF000000"
)

border = Border(
    top=thick,
    left=thick,
    right=thick,
    bottom=thick,
)

no_bottom_border = Border(
    top=thick,
    left=thick,
    right=thick,
)

no_top_border = Border(
    left=thick,
    right=thick,
    bottom=thick,
)

no_tb_border = Border(
    left=thick,
    right=thick,
)

mid_align = Alignment(
    vertical="center",
    horizontal="center",
    wrap_text=True,
)

left_align = Alignment(
    vertical="center",
    horizontal="left",
    wrap_text=True,
)


def merge_intervals(start: int, end: int, ws) -> list:
    _edges = []
    for row in range(start + 1, end + 1):
        if ws[f"A{row}"].value != ws[f"A{start}"].value:
            _edges.append((start, row - 1))
            start = row

    _edges.append((start, end))

    return _edges


def main_styles_acceptation(_edges: list, ws) -> None:
    for start, end in _edges:
        ws.merge_cells(f"A{start}:A{end}")
        ws.merge_cells(f"B{start}:B{end}")
        ws.row_dimensions[2].height = 20

        for row in ws.iter_rows(
                min_row=1,
                max_row=ws.max_row,
                min_col=1,
                max_col=ws.max_column
        ):
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue

                cell.font = font_style
                cell.alignment = mid_align

        for row in ws.iter_rows(
            min_row=1,
            max_row=1,
            min_col=1,
            max_col=ws.max_column
        ):
            for cell in row:
                cell.font = font_style_bold
        for row in ws.iter_rows(
            min_row=2,
            max_row=ws.max_row,
            min_col=ws.max_column,
            max_col=ws.max_column
        ):
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                cell.alignment = left_align

def border_styles(edges: list, ws) -> None:
    for start, end in edges:
        for column in range(ws.min_column, ws.max_column + 1):
            column_letter = get_column_letter(column)
            if isinstance(column_letter, MergedCell):
                continue

            ws[f"{column_letter}{start}"].border = no_bottom_border
            ws[f"{column_letter}{end}"].border = no_top_border
            for number in range(start + 1, end):
                ws[f"{column_letter}{number}"].border = no_tb_border


def add_last_stats(end: int, ws) -> None:
    ws[f"A{end + 1}"] = "Всього: "
    ws[f"B{end + 1}"] = sum(
        row.value for row in ws["B"][1:] if row.value is not None
    )
    ws[f"A{ws.max_row}"].font = font_style_bold
    ws[f"B{ws.max_row}"].font = font_style_bold
    ws[f"A{ws.max_row}"].border = border
    ws[f"B{ws.max_row}"].alignment = mid_align
    ws[f"B{ws.max_row}"].border = border
