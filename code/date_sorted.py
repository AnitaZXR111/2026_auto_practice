import pandas as pd
def date_sort(input_path, output_path):
    df = pd.read_csv(input_path)

    df['日期']=pd.to_datetime(df['日期'])
    df_date = df.sort_values(by="日期",ascending = True).reset_index(drop=True)
    df_date.to_csv(output_path, index=False, encoding='utf-8-sig')

    return(df_date)

if __name__ == "__main__":
    raw_file = "../data/sales_data.csv"
    export_file = "../output/sales_sorted_date.csv"
    data_result = date_sort(raw_file,export_file)
    print(f"已保存至:{export_file}")
