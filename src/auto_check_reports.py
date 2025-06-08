import re
import pandas as pd
from config.regex import WXWORK_FILL_URL
from config.common import CHECK_REPORTS_PATH
from common.validate_util import check_url, check_mission


def read_excel_file(file_path):
    df = pd.read_excel(file_path, header=3)
    df = df.drop(columns=[col for col in df.columns if 'Unnamed' in col])
    return df


def parse_text(text):
    matches = []
    unrecognized_lines = []

    for line in text.strip().split('\n'):
        if line.strip():
            match = re.match(
                WXWORK_FILL_URL, line)

            if match:
                matches.append(match)
            else:
                unrecognized_lines.append(line)
                check_url(line)

    return matches, unrecognized_lines


def run_checks(df,):
    task_text_list = []
    for index, row in df.iterrows():
        project_name = row['项目名称']
        task_name = row['标题']
        task_url = row['链接地址']
        task_time = row['工作时长']
        full_task_text = f'{index}.【{task_time}】{project_name}【{task_name}】{task_url}【{task_time}】'
        task_text_list.append(full_task_text)

    matches, _ = parse_text('\n'.join(task_text_list))
    for match in matches:
        check_mission(match, {'check_date': False})


def main():
    df = read_excel_file(CHECK_REPORTS_PATH)
    run_checks(df)


if __name__ == "__main__":
    main()
