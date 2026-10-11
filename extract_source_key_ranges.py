import json
from openpyxl import load_workbook

sources = [
    (
        r"C:\Users\Admin\Desktop\InviBrowser_GTM_Execution_v1.9_Clean.xlsx",
        {
            "00 Tổng quan": "A1:F32",
            "01 Kế hoạch tuần": "A1:P6",
            "08 Dashboard tiến độ": "A1:N20",
            "03 Lịch nội dung": "A1:S20",
        },
    ),
    (
        r"C:\Users\Admin\Desktop\InviBrowser_Sales_Marketing_Dashboard_v2.0_Clean.xlsx",
        {
            "01 Dashboard": "A1:Q19",
            "02 Funnel Weekly": "A1:O15",
            "07 Sales Performance": "A1:F13",
            "08 Dữ liệu từ Execution": "A1:T5",
        },
    ),
]


def serialize(value):
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


for path, ranges in sources:
    workbook = load_workbook(path, read_only=False, data_only=True)
    print("WORKBOOK", json.dumps(path, ensure_ascii=True))
    for sheet_name, address in ranges.items():
        sheet = workbook[sheet_name]
        matrix = []
        for row in sheet[address]:
            values = [serialize(cell.value) for cell in row]
            if any(value not in (None, "") for value in values):
                matrix.append(values)
        print("SHEET", json.dumps(sheet_name, ensure_ascii=True), "RANGE", address)
        print(json.dumps(matrix, ensure_ascii=True, separators=(",", ":")))
