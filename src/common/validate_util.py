import re
from common.cookie_util import get_mission_info
from common.time_util import convert_iso8601_to_hours, get_week_date_range, date_in_week
from config.common import MISSION_TYPE, HANDLER_NAME
from common.print_util import colored


def check_url(line):
    if '【' not in line or '】' not in line:
        reason = "缺少必要的【】符号"
    elif 'https://' not in line and 'http://' not in line:
        reason = "缺少有效的链接"
    else:
        reason = "格式不符合正则表达式要求"
    print(colored(f"{line}\n{reason}\n", 'yellow'))


def get_estimated_work_hours(task_data):
    desc_html = task_data['description']['html']

    self_estimated_hours = task_data['estimatedTime']
    self_estimated_hours = convert_iso8601_to_hours(self_estimated_hours)

    work_hours_pattern = r'<p class="op-uc-p">预估工时/时长[：:]\s*(\d+\.?\d*).*?</p>'
    work_hours_match = re.search(work_hours_pattern, desc_html, re.S)
    estimated_work_hours_str = work_hours_match.group(
        1).strip() if work_hours_match else ''

    try:
        estimated_work_hours = float(estimated_work_hours_str)
    except ValueError:
        estimated_work_hours = 0.0

    return estimated_work_hours


def check_mission(task_match, option):
    is_task_valid = True
    current_week_mon, current_week_sun = get_week_date_range()

    # 提取正则匹配的任务基础信息
    line_number = task_match.group(1).strip()
    task_estimated_hours_str = task_match.group(2).strip()
    task_project_name = task_match.group(3).strip()
    task_title = task_match.group(4).strip()
    task_link = task_match.group(5).strip()

    # 从任务链接提取任务ID
    task_link_match = re.search(r'/wp/(\d+)', task_link)
    if task_link_match:
        try:
            task_id_from_link = str(int(task_link_match.group(1).strip()))
        except (ValueError, TypeError):
            task_id_from_link = "链接ID解析失败"

    # 获取任务详情
    task_detail = get_mission_info(task_id_from_link)
    if not task_detail or not isinstance(task_detail, dict):
        print(
            colored(f"任务详情获取失败 | 行号: {line_number} | 链接提取ID: {task_id_from_link}", 'red'))
        return False

    embedded_data = task_detail.get('_embedded', {})  # 兜底空字典
    task_status = embedded_data.get('status', {})       # 兜底空字典
    task_project = embedded_data.get('project', {})     # 兜底空字典
    task_responsible = embedded_data.get('responsible', {})  # 兜底空字典
    task_category = embedded_data.get('type', {})       # 兜底空字典

    # 任务核心基础信息
    actual_task_id = str(task_detail.get('id', "接口未返回ID"))
    actual_task_title = task_detail.get('subject', "接口未返回标题")
    raw_self_estimated_hours = task_detail.get('estimatedTime', "")

    # 业务维度详情信息
    task_status_name = task_status.get('name', "状态名称未知")
    actual_project_name = task_project.get('name', "项目名称未知")
    actual_responsible_name = task_responsible.get('name', "处理人未知")
    task_start_date = task_detail.get('startDate', "开始日期未知")
    task_due_date = task_detail.get('dueDate', "截止日期未知")
    task_category_name = task_category.get('name', "任务类型未知")

    print(f"任务类型   ：{task_category_name}")
    print(f"行号       ：{line_number}")
    print(f"项目名称   ：{task_project_name}")
    print(f"任务标题   ：{task_title}")
    print(f"任务ID     ：{actual_task_id}")

    # 校验任务ID一致性
    if task_id_from_link != actual_task_id:
        print(
            colored(f"任务ID不一致 | 预期: {task_id_from_link} | 实际: {actual_task_id}", 'red'))
        is_task_valid = False

    # 校验任务标题一致性
    if actual_task_title != task_title:
        print(
            colored(f"任务标题不一致 | 预期: {task_title} | 实际: {actual_task_title}", 'red'))
        is_task_valid = False

    # 校验任务状态合法性
    if task_status_name not in MISSION_TYPE:
        print(
            colored(f"任务状态异常 | 预期: {task_status_name} | 实际: {MISSION_TYPE}", 'red'))
        is_task_valid = False

    # 校验项目名称一致性
    if actual_project_name != task_project_name:
        print(
            colored(f"项目名称不一致 | 预期: {task_project_name} | 实际: {actual_project_name}", 'red'))
        is_task_valid = False

    # 校验自评时长一致性
    converted_self_estimated_hours = convert_iso8601_to_hours(
        raw_self_estimated_hours)
    if converted_self_estimated_hours != float(task_estimated_hours_str):
        print(
            colored(f"自评时长不一致 | 预期: {task_estimated_hours_str}h | 实际: {converted_self_estimated_hours}h", 'red'))
        is_task_valid = False

    # 校验处理人一致性
    if actual_responsible_name != HANDLER_NAME:
        print(
            colored(f"处理人不一致 | 预期: {HANDLER_NAME} | 实际: {actual_responsible_name}", 'red'))
        is_task_valid = False

    # 校验任务起止日期是否在本周范围内
    is_start_date_legal = date_in_week(
        task_start_date, current_week_mon, current_week_sun)
    is_due_date_legal = date_in_week(
        task_due_date, current_week_mon, current_week_sun)
    # 传入校验日期选项，默认校验
    check_date = option.get('check_date', True)
    if check_date and not (is_start_date_legal and is_due_date_legal):
        print(
            colored(f"超出本周范围 | 开始日期: {task_start_date} | 截止日期: {task_due_date}", 'red'))
        is_task_valid = False

    # 校验预估工时是否正常
    task_estimated_work_hours = get_estimated_work_hours(task_detail)
    if task_estimated_work_hours <= 0:
        print(colored(f"预估工时异常 | 实际: {task_estimated_work_hours}h", 'red'))
        is_task_valid = False

    if is_task_valid:
        print(colored("校验通过", 'green'))

    print("=" * 80, "\n")

    return is_task_valid
