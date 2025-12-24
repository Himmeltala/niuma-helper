import re
import xlwings as xw
import datetime
import requests

TEMPLATE_PATH = rf''
TARGET_TEMPLATE_DIR_PATH = rf'G:\Media Files\Desktop\工作报告'
PATTERN = r'(\d+)\.【([\d.]+)】([^【]+?)【((?:[^】]|【[^】]*】)*?)】(https?://[^\s】]+)【([\d.]+)】'
TEXT = """
"""


def parse_text(text, pattern):
    """
    解析文本内容，识别匹配和未匹配的行，并给出未匹配原因

    :param text: 待解析的文本
    :param pattern: 用于匹配的正则表达式

    :return: 匹配结果列表和未识别行列表
    """
    matches = []
    unrecognized_lines = []

    for line in text.strip().split('\n'):
        # 去除当前行的所有空格
        line = line.replace(" ", "")
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


def convert_duration_to_hours(duration_str: str) -> float:
    """
    将ISO 8601时间间隔格式（PTxxHxxM）转换为以小时为单位的小数

    :params duration_str: 时间间隔字符串，如 "PT3H30M"、"PT1H"、"PT30M"、"PT0M"

    :returns:
        小时数（小数），如 3.5、1.0、0.5、0.0
    """

    # 校验格式合法性，非PT开头/空值直接返回0.0
    if not duration_str or not duration_str.startswith('PT'):
        return 0.0
    # 提取小时数（匹配数字+H的组合，没有则为0）
    hour_match = re.search(r'(\d+)H', duration_str)
    hours = int(hour_match.group(1)) if hour_match else 0
    # 提取分钟数（匹配数字+M且M后面不是S，避免和秒的M混淆，没有则为0）
    minute_match = re.search(r'(\d+)M(?!S)', duration_str)
    minutes = int(minute_match.group(1)) if minute_match else 0
    # 分钟转小时（除以60）后和小时数相加
    total_hours = hours + minutes / 60
    # 保留2位小数，解决浮点数精度问题（如30/60可能出现0.5000000000000001）
    return round(total_hours, 2)


def fill_excel(matches, excel_path):
    """
    将匹配结果填充到 Excel 文件中

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

        url = f'{mission_id}'
        headers = {
        }

        response = requests.get(url, headers=headers)
        data = response.json()

        html = data['description']['html']

        # 自评时长
        estimated_time = data['estimatedTime']
        estimated_time = convert_duration_to_hours(estimated_time)

        # 预估工时/时长
        estimated_duration = re.search(
            r'<p class="op-uc-p">预估工时/时长：(.*?)</p>', html)
        if estimated_duration:
            estimated_duration = estimated_duration.group(1).strip()
        else:
            estimated_duration = ''

        ws.cells(next_row, 2).value = mission_proname
        ws.cells(next_row, 3).value = mission_id
        ws.cells(next_row, 4).value = mission_title
        ws.cells(next_row, 5).value = mission_link
        ws.cells(next_row, 6).value = '已完成'
        ws.cells(next_row, 7).value = '中'
        # 开始时间
        ws.cells(next_row, 8).value = data['startDate']
        # 结束时间
        ws.cells(next_row, 9).value = data['dueDate']
        ws.cells(next_row, 10).value = ''
        # 预估工时
        try:
            ws.cells(next_row, 11).value = float(estimated_duration)
        except (ValueError, TypeError):
            ws.cells(next_row, 11).value = ''
        # 完成工时
        try:
            ws.cells(next_row, 12).value = float(data['customField1'])
        except (ValueError, TypeError):
            ws.cells(next_row, 12).value = ''
        # 完成时间
        ws.cells(next_row, 13).value = data['dueDate']
        # 完成自评时长时长
        ws.cells(next_row, 14).value = float(mission_comptime)

        print(f"序号: {line_num}, 项目名称: {mission_proname}, 任务标题: {mission_title}")
        next_row += 1

    filled_count = len(matches)
    wb.save()
    wb.close()
    app.quit()
    return filled_count


def get_month_week(date=None):
    """
    计算指定日期是当月的第几周（周一为一周起始）

    :param date: 要计算的日期，默认是当前日期

    :return: 当月的周数（int）
    """
    if date is None:
        date = datetime.datetime.now()

    # 获取当月第一天
    first_day = datetime.datetime(date.year, date.month, 1)
    # 计算当月第一天是周几（0=周一, 6=周日，isoweekday()返回1-7，1=周一）
    first_day_weekday = first_day.isoweekday() - 1

    # 计算当前日期与当月第一天的天数差
    day_diff = (date - first_day).days

    # 计算当月第几周（向上取整）
    month_week = (day_diff + first_day_weekday) // 7 + 1
    return month_week


def create_file_from_template(template_path, ):
    """
    根据模板创建当月当周的周报文件

    :param template_path: 模板文件路径

    :return: 新文件的路径
    """
    try:
        app = xw.App(visible=False)
        wb = app.books.open(template_path)

        # 获取当前月份和当月第几周
        current_month = datetime.datetime.now().month
        current_week = get_month_week()  # 使用自定义函数

        # 构造新文件路径
        new_file_path = f'{TARGET_TEMPLATE_DIR_PATH}/{current_month}月/郑人滏-{current_month}月第{current_week}周周报.xlsx'

        # 保存并关闭文件
        wb.save(new_file_path)
        wb.close()
        return new_file_path
    except Exception as e:
        print(f"创建文件失败：{e}")
        raise
    finally:
        # 确保Excel进程退出
        if 'app' in locals():
            app.quit()


def main():
    new_file_path = create_file_from_template(TEMPLATE_PATH)
    matches, unrecognized_lines = parse_text(TEXT, PATTERN)
    recognized_count = len(matches)
    filled_count = fill_excel(matches, new_file_path)

    print(f"识别的数据条数: {recognized_count}")
    print(f"填充到 Excel 的数据条数: {filled_count}")
    print(f"未识别的数据条数: {len(unrecognized_lines)}")


if __name__ == "__main__":
    main()
