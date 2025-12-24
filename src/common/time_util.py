'''
Author: zhengrenfu
Date: 2025-12-23 10:34:02
LastEditors: zhengrenfu
LastEditTime: 2025-12-23 10:34:27
Description: 
'''
import re


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
