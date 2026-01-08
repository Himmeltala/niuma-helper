import re
from common.cookie_util import get_mission_info
from common.time_util import convert_iso8601_to_hours, get_week_date_range, date_in_week
from config.common import MISSION_TYPE, HANDLER_NAME


def check_url(line):
    if '【' not in line or '】' not in line:
        reason = "缺少必要的【】符号"
    elif 'https://' not in line and 'http://' not in line:
        reason = "缺少有效的链接"
    else:
        reason = "格式不符合正则表达式要求"
    print(f"未识别行: {line}\n原因: {reason}\n")


def check_mission(match):
    is_valid = True

    monday, sunday = get_week_date_range()
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

    data = get_mission_info(mission_id)

    embedded = data['_embedded']
    embedded_status = embedded['status']
    embedded_project = embedded['project']
    embedded_responsible = embedded['responsible']
    subject_id = data['id']
    subject = data['subject']
    estimated_time = data['estimatedTime']

    status_name = embedded_status['name']
    project_name = embedded_project['name']
    handler_name = embedded_responsible['name']

    start_date_str = data['startDate'] 
    due_date_str = data['dueDate']    

    if mission_id != str(subject_id):
        print(f"项目ID不一致: {subject_id} -> {mission_id}")
        is_valid = False

    if subject != mission_title:
        print(f"任务标题不一致: {subject} -> {mission_title}")
        is_valid = False

    if status_name not in MISSION_TYPE:
        print(f"任务状态不正常: {status_name}")
        is_valid = False

    if project_name != mission_proname:
        print(f"项目名称不一致: {project_name} -> {mission_proname}")
        is_valid = False

    cdth = convert_iso8601_to_hours(estimated_time)

    if cdth != float(mission_comptime):
        print(
            f"自评时长不一致: {cdth} -> {mission_comptime}")
        is_valid = False

    if handler_name != HANDLER_NAME:
        print(f"处理人不一致：{HANDLER_NAME} -> {handler_name}")
        is_valid = False

    is_start_date_valid = date_in_week(start_date_str, monday, sunday)
    is_due_date_valid = date_in_week(due_date_str, monday, sunday)

    if not (is_start_date_valid and is_due_date_valid):
        print(f"任务 {subject} 日期超出本周范围")
        is_valid = False

    if is_valid:
        print(
            f"\n序号: {line_num}, 项目名称: {mission_proname}, 任务标题: {mission_title}\n")
    return is_valid
