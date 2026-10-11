from openpyxl import load_workbook

path = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx"
workbook = load_workbook(path, read_only=False, data_only=False)
sheet = workbook["Đối thủ"]
for address in ["A4", "A5", "A6", "A7", "A23", "A24", "A25", "A26", "A32"]:
    cell = sheet[address]
    fg = cell.fill.fgColor
    print(address, "style", cell.style_id, "fill", fg.type, fg.rgb, fg.indexed, fg.theme, "font", cell.font.name, cell.font.sz, cell.font.bold, cell.font.color.type if cell.font.color else None, cell.font.color.rgb if cell.font.color and cell.font.color.type == "rgb" else None, "align", cell.alignment.horizontal, cell.alignment.vertical, cell.alignment.wrap_text)
