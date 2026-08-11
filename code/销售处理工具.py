import os
import pandas as pd
import sys
import glob
from datetime import datetime

input_path = "../final_test"
output_path = "../final_output"


def load_data():
    filename = input("请输入Excel文件名：")
    sheetname = input("请输入表单标题：")
    try:
        df = pd.read_excel(os.path.join(input_path, f"{filename}.xlsx"), sheet_name = sheetname)
        df2 = df.copy()
        if '日期' in df2.columns:
            df2['日期'] = pd.to_datetime(df2['日期']).dt.date
        print('加载成功')
        return(df2)
    except:
        print("文件加载失败,请检查文件名是否正确")
        return None

def timestamp():
    ts= datetime.now().strftime("%Y%m%d_%H%M%S")
    return ts

def show_menu():
    print("="*50)
    print('            销售处理工具')
    print("="*50)
    print("1、按销量排序导出")
    print("2、筛选低绩效产品(销量<5)")
    print("3、生成区域统计报表")
    print("4、批量合并多个Excel")
    print("5、退出程序")
    print("="*50)

def option_1():
    df = load_data()
    if df is None:
        return
    
    if '销量' in df.columns:
        sorted_sales = df.sort_values(by = '销量', ascending = False)
        sorted_sales.to_excel(os.path.join(output_path, f"销量排序_{timestamp()}.xlsx"),index = False)
        print(f"已按销量排序，并保存至{output_path}")
    else:
        print("未查找到销量列，请检查源文件。")
        return
    
def option_2():
    df = load_data()
    if df is None: 
        return
    
    if '销量' in df.columns:
        sales_5 = df[df['销量']<5]
        sales_5.to_excel(os.path.join(output_path, f"销量筛选_{timestamp()}.xlsx"),index = False)
        print(f"已筛选出低绩效产品，并保存至{output_path}")
    else:
        print("未查找到销量列，请检查源文件。")
        return

def option_3():
    df = load_data()
    if df is None: 
        return
   
    if '区域' in df.columns:
        if '销量' in df.columns:
            report_sales = df.groupby('区域')['销量'].agg(['count', 'sum', 'mean', 'max', 'min']).reset_index()
            print("区域统计报表：")
            print(report_sales)
            report_sales.to_excel(os.path.join(output_path, f"区域报表_{timestamp()}.xlsx"), index=False)
            print(f"已保存至{output_path}")
        else:
            print("未查找到销量列，请检查源文件。")
            return
    else:
        print("未查找到区域列，请检查源文件。")
        return
        
def option_4():
    excel_files = glob.glob(os.path.join(input_path, "divide/*.xlsx"))
    
    if not excel_files:
        print("当前目录没有Excel文件")
        return

    print(f"\n找到{len(excel_files)}个Excel文件。")
    all_data=[]
    for file in excel_files:
        try:
            df = pd.read_excel(file) 
            df2 = df.copy()
            df2['日期'] = pd.to_datetime(df2['日期']).dt.date
            all_data.append(df2)
            print(f"已加载：{file}")
        except:
            print(f"加载{file}文件时出现错误，请检查源文件。")
        
    if all_data:
        merged = pd.concat(all_data, ignore_index=True)
        merged.to_excel(os.path.join(output_path, f"批量合并_{timestamp()}.xlsx"),index = False)
        print(f"已将文件批量合并，并保存至{output_path}")


def option_5():
    print("\n 感谢使用，再见！")
    sys.exit() 

def main():
    os.makedirs(input_path, exist_ok=True)
    os.makedirs(output_path, exist_ok=True)
    
    while True:
        show_menu()
        choice = input("请选择您需要的业务（1-5）: ")       
        if choice == '1':
            option_1()
        elif choice == '2':
            option_2()
        elif choice == '3':
            option_3()
        elif choice == '4':
            option_4()
        elif choice == '5':
            option_5()
        else:
            print("请输入1-5的数字!")
        input("\n按回车键继续-->")
        
if __name__ == "__main__":
    main()




    
    
    
