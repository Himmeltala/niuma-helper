import pandas as pd
import re
import os

EXCEL_PATH = rf''


def get_file_path(file_path):
    """
    对传入的文件路径进行拆分，获取文件名称（不含后缀）和文件父路径

    :param file_path: 完整的文件路径
    :return: 文件名称（不含后缀）和文件父路径
    """
    parent_path = os.path.dirname(file_path)
    file_name_with_ext = os.path.basename(file_path)
    file_name = os.path.splitext(file_name_with_ext)[0]
    return file_name, parent_path


def read_excel_file(file_path):
    """
    读取 Excel 文件并进行初步处理

    :param file_path: 文件的完整路径
    :return: 处理后的 DataFrame
    """
    df = pd.read_excel(file_path, header=3)
    df = df.drop(columns=[col for col in df.columns if 'Unnamed' in col])
    return df


def check_id_url(row):
    """
    校验 ID 和链接地址是否相等

    :param row: DataFrame 的一行数据
    :return: 布尔值，表示校验结果
    """
    url = row['链接地址']
    if isinstance(url, str):
        last_number = re.findall(r'\d+$', url)
        if last_number:
            last_number = int(last_number[0])
            row_id_number = int(row['ID'])
            return row_id_number == last_number
        else:
            return False
    else:
        return False


def check_dates(row):
    """
    校验预计、结束和完成时间是否都一致

    :param row: DataFrame 的一行数据
    :return: 布尔值，表示校验结果
    """
    return row['预计开始'] == row['预计结束'] == row['完成时间']


def check_hours(row):
    """
    校验预估和完成工时是否一致

    :param row: DataFrame 的一行数据
    :return: 布尔值，表示校验结果
    """
    return row['预估工时'] == row['完成工时']


def check_duplicates(df, columns=["ID", "标题", "链接地址"]):
    """
    校验标题、链接地址是否存在相同的情况

    :param df: 输入的 DataFrame
    :param columns: 要检查重复值的列名列表
    :return: 包含重复值的 DataFrame
    """
    mask = pd.Series([False] * len(df), index=df.index)
    for col in columns:
        mask |= df.duplicated(subset=[col], keep=False)
    return df[mask]


def check_numeric_hours(row):
    """
    校验预估工时、完成工时和工作时长是否都是纯数字

    :param row: DataFrame 的一行数据
    :return: 布尔值，表示校验结果
    """
    # 需要校验的工时列名
    hour_columns = ['预估工时', '完成工时', '工作时长']

    for col in hour_columns:
        # 检查列是否存在
        if col not in row:
            return False

        value = row[col]
        # 检查是否为数字类型
        if not isinstance(value, (int, float)):
            # 尝试转换字符串为数字
            try:
                float(value)
            except (ValueError, TypeError):
                return False
    return True


def run_checks(df, file_name, file_path):
    """
    执行所有校验，并将结果保存到新的 Excel 文件

    :param df: 输入的 DataFrame
    :param file_name: 原文件名
    :param file_path: 文件所在目录路径
    """
    df['ID和链接地址'] = df.apply(check_id_url, axis=1)
    df['预计、结束和完成时间'] = df.apply(check_dates, axis=1)
    df['预估和完成工时'] = df.apply(check_hours, axis=1)
    df['是否纯数字'] = df.apply(check_numeric_hours, axis=1)
    duplicates = check_duplicates(df)

    if not duplicates.empty:
        print("发现重复任务:\n")
        print(duplicates)

    output_file_fullpath = os.path.join(file_path, f"{file_name}_校验结果.xlsx")

    with pd.ExcelWriter(output_file_fullpath, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # 获取工作簿和工作表对象
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # 定义日期格式
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

        # 设置日期列的格式（根据实际列名调整）
        date_columns = ['预计开始', '预计结束', '完成时间']
        for col_name in date_columns:
            if col_name in df.columns:
                # 获取列索引（+1是因为Excel列从1开始）
                col_idx = df.columns.get_loc(col_name) + 1
                # 设置整列格式
                worksheet.set_column(col_idx, col_idx, None, date_format)


def main():
    file_name, parent_path = get_file_path(EXCEL_PATH)
    df = read_excel_file(EXCEL_PATH)
    run_checks(df, file_name, parent_path)


if __name__ == "__main__":
    main()
