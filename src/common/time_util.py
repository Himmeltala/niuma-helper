import re
import datetime


def convert_iso8601_to_hours(duration_str: str) -> float:
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


def get_week_date_range():
    """
    获取本周的日期范围（周一为起始，周日为结束）
    :return: (本周一datetime对象, 本周日datetime对象)
    """
    today = datetime.datetime.now().date()
    # 计算本周一（0=周一，6=周日）
    monday = today - datetime.timedelta(days=today.weekday())
    # 计算本周日
    sunday = monday + datetime.timedelta(days=6)
    return monday, sunday


def date_in_week(date_str, monday, sunday):
    """
    校验日期是否在本周范围内
    :param date_str: 日期字符串（格式：YYYY-MM-DD）
    :param monday: 本周一date对象
    :param sunday: 本周日date对象
    :return: bool - True=在本周，False=超出本周
    """
    try:
        # 解析日期字符串（支持YYYY-MM-DD格式）
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        # 校验是否在本周范围内
        if not (monday <= date_obj <= sunday):
            return False
        return True
    except ValueError as e:
        print(f"无法解析：{e}")
        return False
