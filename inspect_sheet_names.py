from openpyxl import load_workbook

paths = [
    r"C:\Users\Admin\Desktop\Bussiness Plan mẫu.xlsx",
    r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx",
]

for path in paths:
    workbook = load_workbook(path, read_only=True, data_only=False)
    print(path.encode("unicode_escape").decode())
    for index, sheet in enumerate(workbook.worksheets, 1):
        title = sheet.title.encode("unicode_escape").decode()
        print(index, title, sheet.sheet_state)
