import re
from common.validate_util import check_url, check_mission
from config.regex import WXWORK_FILL_URL
from common.print_util import colored

TEXT = """
"""


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


def main():
    matches, _ = parse_text(TEXT)
    total_hours = 0.0  # 初始化累计时长
    has_exception = False  # 标记是否有任务异常

    for match in matches:
        hours_str = match.group(2).strip()
        hours = float(hours_str)
        total_hours += hours
        if not check_mission(match, {'check_date': False}):
            has_exception = True

    print(colored(f"\n已计算时长: {total_hours}小时", 'green'))

    if has_exception:
        print(colored("注意：总时长可能与实际不符！", 'yellow'))


if __name__ == "__main__":
    main()
