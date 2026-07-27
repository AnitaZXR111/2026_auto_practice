import openpyxl

def read_xlsx_value(input_path, sheet, loc):
    wb = openpyxl.load_workbook(input_path)
    df = wb[sheet]
    return df[loc].value

if __name__=="__main__":
    workbook_path = "../data/sales_data.xlsx"
    sheet_n = "sales_data"
    location = "B6"
    target = read_xlsx_value(workbook_path,sheet_n,location)
    print(f"表格{location}位置的值为：{target}")

