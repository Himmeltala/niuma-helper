"""
@description 填充周报或月报
"""

import re
import xlwings as xw

EXCEL_PATH = r'../assets/targets/6月第1周报.xlsx'
PATTERN = r'(\d+)\.【([\d.]+)】([^【]+?)【((?:[^】]|【[^】]*】)*?)】(https?://[^\s】]+)【([\d.]+)】'
TEXT = """
"""


def parse_text(text, pattern):
    """
    解析文本内容，识别匹配和未匹配的行，并给出未匹配原因。

    :param text: 待解析的文本
    :param pattern: 用于匹配的正则表达式
    :return: 匹配结果列表和未识别行列表
    """
    matches = []
    unrecognized_lines = []

    for line in text.strip().split('\n'):
        if line.strip():
            match = re.match(pattern, line)
            if match:
                matches.append(match)
            else:
                unrecognized_lines.append(line)
                if '【' not in line or '】' not in line:
                    reason = "缺少必要的【】符号"
                elif 'https://' not in line and 'http://' not in line:
                    reason = "缺少有效的链接"
                else:
                    reason = "格式不符合正则表达式要求"
                print(f"未识别行: {line}\n原因: {reason}\n")
    return matches, unrecognized_lines


def fill_excel(matches, excel_path):
    """
    将匹配结果填充到 Excel 文件中。

    :param matches: 匹配结果列表
    :param excel_path: Excel 文件路径
    :return: 填充的数据条数
    """
    next_row = 5
    app = xw.App(visible=False)
    wb = app.books.open(excel_path)
    ws = wb.sheets[0]

    for match in matches:
        line_num = match.group(1).strip()
        mission_comptime = match.group(2).strip()
        mission_proname = match.group(3).strip()
        mission_title = match.group(4).strip()
        mission_link = match.group(5).strip()

        mission_link_mt = re.search(r'/wp/(\d+)', mission_link)
        mission_id = ''
        if mission_link_mt:
            mission_id = mission_link_mt.group(1)
        mission_id = str(int(mission_id))

        print(f"序号: {line_num}, 项目名称: {mission_proname}, 任务标题: {mission_title}")

        ws.cells(next_row, 2).value = mission_proname
        ws.cells(next_row, 3).value = mission_id
        ws.cells(next_row, 4).value = mission_title
        ws.cells(next_row, 5).value = mission_link
        ws.cells(next_row, 6).value = '已完成'
        ws.cells(next_row, 7).value = '中'
        ws.cells(next_row, 14).value = mission_comptime
        next_row += 1

    filled_count = len(matches)
    wb.save()
    wb.close()
    app.quit()
    return filled_count


def main():
    matches, unrecognized_lines = parse_text(TEXT, PATTERN)
    recognized_count = len(matches)
    filled_count = fill_excel(matches, EXCEL_PATH)

    print(f"识别的数据条数: {recognized_count}")
    print(f"填充到 Excel 的数据条数: {filled_count}")
    print(f"未识别的数据条数: {len(unrecognized_lines)}")


if __name__ == "__main__":
    main()
