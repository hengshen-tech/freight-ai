"""生成演示用的样例单证和报关单模板（虚构数据）。运行：python scripts/make_samples.py"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = Path(__file__).resolve().parent.parent / "samples"
OUT.mkdir(exist_ok=True)
F = Font(name="Arial", size=10); B = Font(name="Arial", size=10, bold=True); T = Font(name="Arial", size=14, bold=True)
thin = Side(style="thin", color="999999"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD = PatternFill("solid", fgColor="E8EEF4"); INPUT = PatternFill("solid", fgColor="FFF9DB")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="center", wrap_text=True)

SELLER = "NINGBO HENGYI ELECTRIC APPLIANCE CO., LTD.\nNo.88 Demo Road, Beilun District, Ningbo, Zhejiang, China"
SELLER_CN = "宁波恒益电器有限公司（演示数据）"
BUYER = "NORDLICHT HANDELS GMBH\nBeispielstrasse 12, 20457 Hamburg, Germany"
ITEMS = [  # 品名, 中文, 型号, HS, 数量, 单位, 单价, 箱数, 净重, 毛重
    ("LED PANEL LIGHT 600x600 40W", "LED面板灯", "HY-PL6060-40W", "9405119000", 1200, "PCS", 8.50, 150, 2880.0, 3150.0),
    ("LED DOWNLIGHT 6 INCH 15W", "LED筒灯", "HY-DL6-15W", "9405119000", 3000, "PCS", 3.20, 125, 1350.0, 1525.0),
    ("LED DRIVER 40W CONSTANT CURRENT", "LED驱动电源", "HY-DR40", "8504401400", 1500, "PCS", 2.10, 30, 420.0, 462.0),
]

def setw(ws, widths):
    for col, w in zip("ABCDEFGHIJ", widths): ws.column_dimensions[col].width = w

def kv(ws, row, label, value, span="B{r}:D{r}"):
    ws[f"A{row}"] = label; ws[f"A{row}"].font = B
    ws[f"B{row}"] = value; ws[f"B{row}"].font = F; ws[f"B{row}"].alignment = WRAP
    ws.merge_cells(span.format(r=row))

def table_head(ws, row, heads):
    for i, h in enumerate(heads):
        c = ws.cell(row, i + 1, h); c.font = B; c.fill = HEAD; c.border = BOX; c.alignment = CENTER

def invoice():
    wb = Workbook(); ws = wb.active; ws.title = "INVOICE"; setw(ws, [16, 34, 18, 14, 10, 8, 12, 14])
    ws["A1"] = "COMMERCIAL INVOICE"; ws["A1"].font = T; ws.merge_cells("A1:H1"); ws["A1"].alignment = CENTER
    kv(ws, 3, "Seller:", SELLER); ws.row_dimensions[3].height = 30
    kv(ws, 4, "Buyer:", BUYER); ws.row_dimensions[4].height = 30
    ws["F3"] = "Invoice No.:"; ws["G3"] = "HY-INV-261103"; ws["F4"] = "Date:"; ws["G4"] = "2026-11-03"
    kv(ws, 5, "Contract No.:", "HY-SC-2026-0917"); kv(ws, 6, "Terms:", "FOB NINGBO"); kv(ws, 7, "Payment:", "T/T 30 DAYS")
    kv(ws, 8, "From / To:", "NINGBO, CHINA  →  HAMBURG, GERMANY"); kv(ws, 9, "Currency:", "USD")
    for c in ("F3", "F4"): ws[c].font = B
    table_head(ws, 11, ["No.", "Description", "Model", "HS Code", "Qty", "Unit", "Unit Price", "Amount"])
    r = 12
    for i, it in enumerate(ITEMS, 1):
        for j, v in enumerate([i, it[0], it[2], it[3], it[4], it[5], it[6], f"=E{r}*G{r}"]):
            c = ws.cell(r, j + 1, v); c.font = F; c.border = BOX
        ws.cell(r, 7).number_format = ws.cell(r, 8).number_format = "#,##0.00"; r += 1
    ws[f"A{r}"] = "TOTAL"; ws[f"A{r}"].font = B; ws[f"E{r}"] = f"=SUM(E12:E{r-1})"; ws[f"H{r}"] = f"=SUM(H12:H{r-1})"
    ws[f"H{r}"].number_format = "#,##0.00"; ws[f"H{r}"].font = B
    ws[f"A{r+2}"] = "Country of Origin: CHINA"; ws[f"A{r+3}"] = "Shipping Marks: NORDLICHT / HAMBURG / C/NO.1-305"
    wb.save(OUT / "样例_商业发票.xlsx")

def packing():
    wb = Workbook(); ws = wb.active; ws.title = "PACKING LIST"; setw(ws, [8, 34, 18, 10, 8, 10, 12, 12, 12])
    ws["A1"] = "PACKING LIST"; ws["A1"].font = T; ws.merge_cells("A1:I1"); ws["A1"].alignment = CENTER
    kv(ws, 3, "Shipper:", SELLER, "B{r}:E{r}"); ws.row_dimensions[3].height = 30
    kv(ws, 4, "Consignee:", BUYER, "B{r}:E{r}"); ws.row_dimensions[4].height = 30
    kv(ws, 5, "Invoice No.:", "HY-INV-261103", "B{r}:E{r}"); kv(ws, 6, "Marks:", "NORDLICHT / HAMBURG / C/NO.1-305", "B{r}:E{r}")
    table_head(ws, 8, ["No.", "Description", "Model", "Qty", "Unit", "Cartons", "N.W.(KG)", "G.W.(KG)", "CBM"])
    r = 9
    for i, it in enumerate(ITEMS, 1):
        for j, v in enumerate([i, it[0], it[2], it[4], it[5], it[7], it[8], it[9], round(it[7] * 0.075, 2)]):
            c = ws.cell(r, j + 1, v); c.font = F; c.border = BOX
        r += 1
    ws[f"A{r}"] = "TOTAL"; ws[f"A{r}"].font = B
    for col in "DFGHI": ws[f"{col}{r}"] = f"=SUM({col}9:{col}{r-1})"; ws[f"{col}{r}"].font = B
    ws[f"A{r+2}"] = "Packing: EXPORT CARTONS"; ws[f"A{r+3}"] = "Container: 1 x 40'HQ"
    wb.save(OUT / "样例_装箱单.xlsx")

def contract():
    wb = Workbook(); ws = wb.active; ws.title = "合同"; setw(ws, [18, 40, 16, 14, 14])
    ws["A1"] = "销售合同 SALES CONTRACT"; ws["A1"].font = T; ws.merge_cells("A1:E1"); ws["A1"].alignment = CENTER
    rows = [("合同号 Contract No.", "HY-SC-2026-0917"), ("签订日期 Date", "2026-09-17"),
            ("卖方 Seller", SELLER_CN + "\n" + SELLER), ("统一社会信用代码", "91330206DEMO00000X"),
            ("买方 Buyer", BUYER), ("价格条款 Price Terms", "FOB NINGBO"), ("币种 Currency", "USD"),
            ("合同总金额 Total Amount", 22950.00), ("装运港 Port of Loading", "NINGBO, CHINA"),
            ("目的港 Port of Destination", "HAMBURG, GERMANY"), ("运输方式 Shipment", "BY SEA, 1x40'HQ"),
            ("订舱信息 Booking", "Vessel/Voy: DEMO HARMONY V.046W   B/L No.: DEMO2611003"),
            ("生产厂所在地", "浙江省宁波市")]
    for i, (k, v) in enumerate(rows, 3):
        kv(ws, i, k, v, "B{r}:E{r}")
        if "\n" in str(v): ws.row_dimensions[i].height = 45
    wb.save(OUT / "样例_销售合同.xlsx")

def template():
    wb = Workbook(); ws = wb.active; ws.title = "报关单"; setw(ws, [7, 16, 30, 14, 12, 14, 10, 14, 14, 12])
    ws["A1"] = "中华人民共和国海关出口货物报关单（草单）"; ws["A1"].font = T; ws.merge_cells("A1:J1"); ws["A1"].alignment = CENTER
    ws["A2"] = "说明：黄色单元格为待填项，其余请勿改动。"; ws["A2"].font = Font(name="Arial", size=9, italic=True, color="666666"); ws.merge_cells("A2:J2")
    head = [  # 每行两组：标签(A:B) 值(C:E) | 标签(F:G) 值(H:J)
        ("境内发货人", "出境关别"), ("统一社会信用代码", "出口日期"), ("境外收货人", "申报日期"),
        ("生产销售单位", "备案号"), ("运输方式", "运输工具名称及航次号"), ("提运单号", "合同协议号"),
        ("监管方式", "征免性质"), ("贸易国（地区）", "运抵国（地区）"), ("指运港", "离境口岸"),
        ("成交方式", "许可证号"), ("运费", "保费"), ("包装种类", "件数"), ("毛重（千克）", "净重（千克）"),
        ("随附单证", "境内货源地"),
    ]
    r = 3
    for a, b in head:
        for lab, lc, vc, lm, vm in ((a, "A", "C", f"A{r}:B{r}", f"C{r}:E{r}"), (b, "F", "H", f"F{r}:G{r}", f"H{r}:J{r}")):
            ws[f"{lc}{r}"] = lab; ws[f"{lc}{r}"].font = B; ws[f"{lc}{r}"].fill = HEAD; ws[f"{lc}{r}"].alignment = WRAP
            ws.merge_cells(lm); ws.merge_cells(vm); ws[f"{vc}{r}"].fill = INPUT; ws[f"{vc}{r}"].font = F; ws[f"{vc}{r}"].alignment = WRAP
        for col in "ABCDEFGHIJ": ws[f"{col}{r}"].border = BOX
        r += 1
    ws[f"A{r}"] = "标记唛码及备注"; ws[f"A{r}"].font = B; ws[f"A{r}"].fill = HEAD; ws.merge_cells(f"A{r}:B{r}")
    ws.merge_cells(f"C{r}:J{r}"); ws[f"C{r}"].fill = INPUT; ws[f"C{r}"].alignment = WRAP; ws.row_dimensions[r].height = 32
    for col in "ABCDEFGHIJ": ws[f"{col}{r}"].border = BOX
    r += 2
    table_head(ws, r, ["项号", "商品编号", "商品名称及规格型号", "数量", "单位", "单价", "币制", "总价", "原产国（地区）", "最终目的国"])
    first = r + 1
    for i in range(first, first + 8):
        for col in "ABCDEFGHIJ":
            c = ws[f"{col}{i}"]; c.border = BOX; c.font = F; c.alignment = WRAP
            if col != "H": c.fill = INPUT
        ws[f"H{i}"] = f'=IF(AND(D{i}<>"",F{i}<>""),D{i}*F{i},"")'; ws[f"H{i}"].number_format = "#,##0.00"
        ws[f"F{i}"].number_format = "#,##0.0000"
    last = first + 7; r = last + 1
    ws[f"A{r}"] = "合计"; ws[f"A{r}"].font = B; ws.merge_cells(f"A{r}:C{r}")
    ws[f"D{r}"] = f"=SUM(D{first}:D{last})"; ws[f"H{r}"] = f"=SUM(H{first}:H{last})"; ws[f"H{r}"].number_format = "#,##0.00"
    for col in "ABCDEFGHIJ": ws[f"{col}{r}"].border = BOX; ws[f"{col}{r}"].font = B
    ws[f"A{r+2}"] = "报关人员：              申报单位（签章）：              海关批注及签章："; ws[f"A{r+2}"].font = F
    wb.save(OUT / "报关单模板.xlsx")

if __name__ == "__main__":
    invoice(); packing(); contract(); template(); print("samples written to", OUT)
