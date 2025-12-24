import requests
import re
from common.time_util import convert_iso8601_to_hours


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


def check_mission(matches):
    """
    检查匹配结果是否符合要求

    :param matches: 匹配结果列表

    :return: 无
    """

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

        url = f'{mission_id}'
        headers = {
        }

        response = requests.get(url, headers=headers)
        data = response.json()

        embedded = data['_embedded']
        embedded_status = embedded['status']
        embedded_project = embedded['project']
        subject_id = data['id']
        subject = data['subject']
        html = data['description']['html']
        estimated_time = data['estimatedTime']

        status_name = embedded_status['name']
        project_name = embedded_project['name']

        match = re.search(
            r'<p class="op-uc-p">预估工时/时长：(.*?)</p>', html)

        if mission_id != str(subject_id):
            print(f"项目ID不一致: {mission_id} -> {subject_id}")

        if subject != mission_title:
            print(f"任务标题不一致: {subject} -> {mission_title}")

        if status_name != '待检查':
            print(f"任务状态不是待检查: {status_name}")

        if project_name != mission_proname:
            print(f"项目名称不一致: {project_name} -> {mission_proname}")

        cdth = convert_iso8601_to_hours(estimated_time)

        if cdth != float(mission_comptime):
            print(
                f"自评时长不一致: {cdth} -> {mission_comptime}")


def main():
    matches, unrecognized_lines = parse_text(TEXT, PATTERN)
    check_mission(matches)


if __name__ == "__main__":
    main()
