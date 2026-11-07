"""Build BoQ_Measurement_Template.xlsx: dimension sheet -> bill of quantities -> trade summary,
plus a queries / assumptions register. Run: python build_template.py"""
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

NAVY, LIGHT, GROUP, INPUT = "1F3A5F", "E8EEF5", "D6E2EF", "FFF4CC"
thin = Side(style="thin", color="B7C3D0")
B = Border(left=thin, right=thin, top=thin, bottom=thin)
WHITE = Font(bold=True, color="FFFFFF")
BOLD = Font(bold=True)

# example: one bedroom, finishes trades (fictitious)
BILL = [  # trade, item, description, unit
    ("FLOOR FINISHES", "F.01", "Porcelain floor tiles 600x600 on adhesive, to bedrooms", "m2"),
    ("FLOOR FINISHES", "F.02", "Tile skirting 100 high", "m"),
    ("WALL FINISHES", "W.01", "Emulsion paint, primer + 2 coats, to plastered walls", "m2"),
    ("CEILING FINISHES", "C.01", "Gypsum board suspended ceiling, painted", "m2"),
    ("DOORS", "D.01", "Internal door set 900 x 2100, complete", "nr"),
]
DIMS = [  # item, location, description, drawing, times, no, L, W, H, deduct, note
    ("F.01", "Bedroom 1", "Floor", "A-101", None, 1, 3.60, 3.20, None, "", ""),
    ("F.01", "Bedroom 1", "Wardrobe recess", "A-101", None, 1, 1.80, 0.60, None, "", ""),
    ("F.01", "Bedroom 1", "Ddt wardrobe base unit", "A-101", None, 1, 1.80, 0.60, None, "D", "Tiled under? see Q-02"),
    ("F.02", "Bedroom 1", "Perimeter", "A-101", 2, 1, 3.60 + 3.20, None, None, "", ""),
    ("F.02", "Bedroom 1", "Ddt door opening", "A-101", None, 1, 0.90, None, None, "D", ""),
    ("W.01", "Bedroom 1", "Walls, perimeter x height", "A-101 / A-301", 2, 1, 3.60 + 3.20, None, 2.70, "", "Height to ceiling, see Q-01"),
    ("W.01", "Bedroom 1", "Ddt window", "A-101", None, 1, 1.50, None, 1.50, "D", ""),
    ("W.01", "Bedroom 1", "Ddt door opening", "A-101", None, 1, 0.90, None, 2.10, "D", ""),
    ("C.01", "Bedroom 1", "Ceiling", "A-201", None, 1, 3.60, 3.20, None, "", ""),
    ("D.01", "Bedroom 1", "Door D-07", "A-101 / door schedule", None, 1, None, None, None, "", ""),
]
QUERIES = [
    ("Q-01", "A-301", "Ceiling height not dimensioned in bedrooms", "2.70 m taken from section A-A", "Open"),
    ("Q-02", "A-101", "Floor tiling under fitted wardrobes not stated", "Tiling omitted under base units", "Open"),
]


def head(ws, row, titles, widths=None):
    for c, t in enumerate(titles, start=1):
        cell = ws.cell(row, c, t)
        cell.font, cell.border = WHITE, B
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    if widths:
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(1, c).column_letter].width = w
    ws.row_dimensions[row].height = 30


def build(path="BoQ_Measurement_Template.xlsx"):
    wb = Workbook()

    # Cover
    cv = wb.active
    cv.title = "Cover"
    cv["A1"] = "BILL OF QUANTITIES"
    cv["A1"].font = Font(bold=True, size=16, color=NAVY)
    for i, k in enumerate(["Project", "Client", "Drawings / revision", "Measurement basis",
                           "Prepared by", "Date", "Currency"], start=3):
        cv.cell(i, 1, k).font = BOLD
        cv.cell(i, 2).fill = PatternFill("solid", fgColor=INPUT)
        cv.cell(i, 2).border = B
    cv["B6"] = "e.g. NRM2 / ASMM / POMI, net in place, no waste unless stated"
    cv["A12"] = "How to use"
    cv["A12"].font = Font(bold=True, color=NAVY)
    steps = ["1. Dim Sheet: one line per measurement; Times x No x L x W x H (blank = not used). 'D' in Deduct makes it negative.",
             "2. BoQ: quantities are pulled automatically from the Dim Sheet by item code. Enter rates in the yellow column.",
             "3. Summary: trade totals, contingency and grand total update by themselves.",
             "4. Queries: every assumption made while measuring, with the drawing it relates to.",
             "Yellow cells are inputs. Everything else is formulas - check any number by tracing it back to the Dim Sheet."]
    for i, s in enumerate(steps, start=13):
        cv.cell(i, 1, s)
    cv.column_dimensions["A"].width = 24
    cv.column_dimensions["B"].width = 70

    # Dim sheet
    ds = wb.create_sheet("Dim Sheet")
    ds["A1"] = "DIMENSION SHEET"
    ds["A1"].font = Font(bold=True, size=14, color=NAVY)
    head(ds, 3, ["Item", "Location", "Description of measurement", "Drawing", "Times", "No.",
                 "L (m)", "W (m)", "H (m)", "Deduct (D)", "Quantity", "Unit", "Notes"],
         [8, 14, 38, 16, 7, 6, 8, 8, 8, 9, 11, 6, 30])
    r0 = 4
    nrows = 200
    for i in range(nrows):
        r = r0 + i
        d = DIMS[i] if i < len(DIMS) else None
        vals = list(d[:10]) + [None, None, d[10]] if d else [None] * 13
        for c, v in enumerate(vals, start=1):
            cell = ds.cell(r, c, v)
            cell.border = B
            if c in (5, 6):
                cell.number_format = "0"
            elif 7 <= c <= 9:
                cell.number_format = "0.00"
        ds.cell(r, 11, f'=IF(OR(A{r}="",COUNT(E{r}:I{r})=0),"",PRODUCT(E{r}:I{r})*IF(UPPER(J{r})="D",-1,1))').number_format = "0.00;[Red]-0.00"
        ds.cell(r, 12, f'=IFERROR(INDEX(BoQ!$C:$C,MATCH(A{r},BoQ!$A:$A,0)),"")')
    dv = DataValidation(type="list", formula1='"D"', allow_blank=True)
    ds.add_data_validation(dv)
    dv.add(f"J{r0}:J{r0+nrows-1}")
    ds.conditional_formatting.add(f"A{r0}:M{r0+nrows-1}",
                                  FormulaRule(formula=[f'UPPER($J{r0})="D"'], font=Font(color="C00000")))
    ds.freeze_panes = "A4"

    # BoQ
    bq = wb.create_sheet("BoQ")
    bq["A1"] = "BILL OF QUANTITIES"
    bq["A1"].font = Font(bold=True, size=14, color=NAVY)
    head(bq, 3, ["Item", "Description", "Unit", "Qty", "Rate", "Amount", "Trade"], [8, 58, 7, 11, 11, 14, 4])
    r = 4
    trade_rows = {}
    for trade in dict.fromkeys(t for t, *_ in BILL):
        bq.cell(r, 2, trade).font = BOLD
        for c in range(1, 7):
            bq.cell(r, c).fill = PatternFill("solid", fgColor=GROUP)
            bq.cell(r, c).border = B
        r += 1
        g0 = r
        for t, item, desc, unit in BILL:
            if t != trade:
                continue
            bq.cell(r, 1, item)
            bq.cell(r, 2, desc).alignment = Alignment(wrap_text=True)
            bq.cell(r, 3, unit)
            bq.cell(r, 4, f"=ROUND(SUMIF('Dim Sheet'!$A:$A,A{r},'Dim Sheet'!$K:$K),2)").number_format = "#,##0.00"
            bq.cell(r, 5).fill = PatternFill("solid", fgColor=INPUT)
            bq.cell(r, 5).number_format = "#,##0.00"
            bq.cell(r, 6, f'=IF(E{r}="","",D{r}*E{r})').number_format = "#,##0.00"
            bq.cell(r, 7, trade)
            for c in range(1, 7):
                bq.cell(r, c).border = B
            r += 1
        bq.cell(r, 2, f"Total carried to summary - {trade}").font = BOLD
        bq.cell(r, 6, f"=SUM(F{g0}:F{r-1})").number_format = "#,##0.00"
        for c in range(1, 7):
            bq.cell(r, c).fill = PatternFill("solid", fgColor=LIGHT)
            bq.cell(r, c).border = B
        trade_rows[trade] = r
        r += 2
    bq.column_dimensions["G"].hidden = True
    bq.freeze_panes = "A4"

    # Summary
    sm = wb.create_sheet("Summary")
    sm["A1"] = "SUMMARY"
    sm["A1"].font = Font(bold=True, size=14, color=NAVY)
    head(sm, 3, ["No.", "Trade", "Amount"], [6, 40, 16])
    r = 4
    for i, (trade, tr) in enumerate(trade_rows.items(), start=1):
        sm.cell(r, 1, i)
        sm.cell(r, 2, trade)
        sm.cell(r, 3, f"=BoQ!F{tr}").number_format = "#,##0.00"
        for c in range(1, 4):
            sm.cell(r, c).border = B
        r += 1
    sub = r
    sm.cell(r, 2, "Sub-total").font = BOLD
    sm.cell(r, 3, f"=SUM(C4:C{r-1})").number_format = "#,##0.00"
    sm.cell(r + 1, 2, "Contingency")
    sm.cell(r + 1, 4, 0.05).number_format = "0%"
    sm.cell(r + 1, 4).fill = PatternFill("solid", fgColor=INPUT)
    sm.cell(r + 1, 3, f"=C{sub}*D{sub+1}").number_format = "#,##0.00"
    sm.cell(r + 2, 2, "TOTAL").font = BOLD
    sm.cell(r + 2, 3, f"=C{sub}+C{sub+1}").number_format = "#,##0.00"
    sm.cell(r + 2, 3).font = BOLD
    for rr in range(sub, sub + 3):
        for c in range(1, 4):
            sm.cell(rr, c).fill = PatternFill("solid", fgColor=LIGHT)
            sm.cell(rr, c).border = B

    # Queries
    qs = wb.create_sheet("Queries")
    qs["A1"] = "QUERIES AND ASSUMPTIONS"
    qs["A1"].font = Font(bold=True, size=14, color=NAVY)
    head(qs, 3, ["Ref", "Drawing", "Query / missing information", "Assumption made in the measure", "Status"],
         [7, 12, 48, 44, 10])
    for i, q in enumerate(QUERIES, start=4):
        for c, v in enumerate(q, start=1):
            qs.cell(i, c, v).border = B
            qs.cell(i, c).alignment = Alignment(wrap_text=True, vertical="top")

    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ds.print_area = f"A1:M{r0 + len(DIMS) + 2}"
    wb.save(path)
    print(path)


if __name__ == "__main__":
    build()
