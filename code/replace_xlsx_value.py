import openpyxl

def replace_xlsx_value(input_path, output_path, sheet, loc,new_val):
    wb = openpyxl.load_workbook(input_path)
    df = wb[sheet]

    df[loc]=new_val
    wb.save(output_path)

if __name__=="__main__":
    old_path = "../data/sales_data.xlsx"
    new_path = "../output/sales_data_update.xlsx"
    sheet_n = "sales_data"
    location = "C5"
    new_value = 'A'
    replace_xlsx_value(old_path,new_path,sheet_n,location,new_value)
    print(f"已将数据表{location}的值更新为{new_value},并保存至:{new_path}")
