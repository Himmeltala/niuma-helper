"""
@description 提取文本中我的任务；清理不符合规则的任务；检查日报填写是否规范
"""

import re

PATTERN = r'(\d+)\.([^【]+?)【((?:[^】]|【[^】]*】)*?)】(https?://[^\s】]+)【([\d.]+)】'
TEXT = """
"""


def format_text(text, name=None):
    """
    格式化文本，去除多余的空格和换行符，添加换行符分隔每行，
    并删除不包含指定名字的行。

    :param text: 待格式化的文本
    :param name: 要匹配的名字
    :return: 格式化后的文本
    """
    text = text.replace(" ", "")
    # 名字提供了，删除不包含名字的行
    if name:
        lines = text.strip().split('\n')
        pick_lines = [line for line in lines if name in line]
        text = '\n'.join(pick_lines)
    return text


def parse_title(text):
    """
    从给定文本中提取任务标题，并清理标题内容，去除第一个 【 之前和最后一个 】 之后的内容。

    :param text: 包含任务标题的原始文本
    :return: 清理后的任务标题，如果未找到有效标题则返回空字符串
    """
    title_start = text.find('【')
    http_start = min(
        text.find('https://') if 'https://' in text else len(text),
        text.find('http://') if 'http://' in text else len(text)
    )
    if title_start != -1 and http_start > title_start:
        title_part = text[title_start + 1:http_start].strip()
        title_start_in_title = title_part.find('【')
        title_end_in_title = title_part.rfind('】')
        if title_start_in_title == -1:
            return title_part[:title_end_in_title + 1]
        if title_end_in_title == -1:
            return title_part[title_start_in_title:]
        if title_start_in_title != -1 and title_end_in_title != -1:
            return title_part[title_start_in_title:title_end_in_title + 1]
    else:
        return ''


def parse_text(text, pattern):
    """
    解析文本内容，识别匹配和未匹配的行，并给出未匹配原因。

    :param text: 待解析的文本
    :param pattern: 用于匹配的正则表达式
    :return: 匹配结果列表、未识别行列表和清理前后数据的映射
    """
    matches = []
    unrecognized_lines = []
    cleaned_mapping = {}  # 存储清理前后数据的映射

    text = format_text(text)

    for line in text.strip().split('\n'):
        if line.strip():
            match = re.match(pattern, line)
            if match:
                matches.append(match)
            else:
                # 提取序号
                num_match = re.search(r'(\d+)\.', line)
                num_part = num_match.group(0) if num_match else ''
                # 提取项目名称
                project_name_match = re.search(r'\.([^【]+?)【', line)
                project_name_part = project_name_match.group(
                    1) if project_name_match else ''
                # 提取任务标题
                title_part = parse_title(line)
                # 提取链接
                link_match = re.search(r'(https?://[^【]+)(?=【)', line)
                link_part = link_match.group(1) if link_match else ''
                # 提取数值
                value_match = re.search(r'【([\d.]+)】', line)
                value_part = value_match.group(1) if value_match else ''

                # 重新组合文本
                cleaned_line = f"{num_part}{project_name_part}【{title_part}{link_part}【{value_part}】"

                new_match = re.match(pattern, cleaned_line)
                if new_match:
                    matches.append(new_match)
                    cleaned_mapping[line] = cleaned_line
                    continue

                unrecognized_lines.append(line)

                if '【' not in line or '】' not in line:
                    reason = "缺少必要的【】符号"
                elif 'https://' not in line and 'http://' not in line:
                    reason = "缺少有效的链接"
                else:
                    reason = "格式不符合正则表达式要求"
                print(f"未识别行: {line}\n原因: {reason}\n")
    return matches, unrecognized_lines, cleaned_mapping


def main():
    matches, unrecognized_lines, cleaned_mapping = parse_text(TEXT, PATTERN)
    recognized_count = len(matches)

    if matches:
        title = " 识别的数据 "
        print(f"{'#' * 20}{title}{'#' * 20}")
        for match in matches:
            print(match.group(0))

    if unrecognized_lines:
        title = " 未识别的数据 "
        print(f"{'#' * 20}{title}{'#' * 20}")
        for line in unrecognized_lines:
            print(line)

    if cleaned_mapping:
        title = " 清理前后的数据 "
        print(f"\n{'#' * 20}{title}{'#' * 20}")
        for original, cleaned in cleaned_mapping.items():
            print(f"清理前: {original}")
            print(f"清理后: {cleaned}")

    print(f"\n识别的数据条数: {recognized_count}")
    print(f"未识别的数据条数: {len(unrecognized_lines)}")


if __name__ == "__main__":
    main()
