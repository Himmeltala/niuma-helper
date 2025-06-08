import re
import datetime
import os
import xlwings as xw
from common.cookie_util import get_mission_info
from common.time_util import get_month_week
from common.validate_util import check_url, check_mission
from config.common import HANDLER_NAME, TEMPLATE_PATH, TARGET_TEMPLATE_DIR_PATH
from config.regex import WXWORK_FILL_URL

TEXT = """
"""

FILL_START_ROW = 5


def parse_task_text(text_content, match_pattern):
    """
    解析任务文本内容，识别匹配和未匹配的行数据
    :param text_content: 待解析的原始文本
    :param match_pattern: 用于匹配任务的正则表达式
    :return: 匹配成功的任务正则结果列表、未识别的文本行列表
    """
    matched_task_list = []
    unrecognized_line_list = []

    # 遍历每行文本做匹配处理
    for single_line in text_content.strip().split('\n'):
        trim_line = single_line.strip()
        if trim_line:
            line_match = re.match(match_pattern, trim_line)
            if line_match:
                matched_task_list.append(line_match)
            else:
                unrecognized_line_list.append(trim_line)
                check_url(trim_line)

    return matched_task_list, unrecognized_line_list


def fill_task_data_to_excel(matched_task_list, excel_file_path):
    """
    xlwings操作Excel，将校验通过的任务数据填充至表格中
    :param matched_task_list: 正则匹配成功的任务结果列表
    :param excel_file_path: Excel文件的完整路径
    :return: 成功填充的数据条数
    """
    current_fill_row = FILL_START_ROW
    excel_app = None
    excel_workbook = None
    excel_worksheet = None
    success_fill_count = 0

    try:
        excel_app = xw.App(visible=False, add_book=False)
        excel_workbook = excel_app.books.open(excel_file_path, read_only=False)
        excel_worksheet = excel_workbook.sheets[0]

        # 遍历匹配的任务，校验通过则填充数据
        for task_match in matched_task_list:
            check_mission(task_match, {'check_date': False})

            # 提取正则匹配的任务基础数据
            task_serial_num = task_match.group(1).strip()
            task_self_est_hours_str = task_match.group(2).strip()
            task_project_name = task_match.group(3).strip()
            task_title = task_match.group(4).strip()
            task_url = task_match.group(5).strip()

            task_id_match = re.search(r'/wp/(\d+)', task_url)
            task_id = ""
            if task_id_match:
                try:
                    task_id = str(int(task_id_match.group(1).strip()))
                except (ValueError, TypeError):
                    task_id = ""

            if not task_id:
                continue

            # 获取任务详情完整数据
            task_detail_data = get_mission_info(task_id)
            task_desc_html = task_detail_data['description']['html']

            # 解析任务预估工时
            est_work_hours_match = re.search(
                r'<p class="op-uc-p">预估工时/时长[：:]\s*(\d+\.?\d*).*?</p>',
                task_desc_html,
                re.S
            )
            est_work_hours_str = est_work_hours_match.group(
                1).strip() if est_work_hours_match else ''

            # 项目名称
            excel_worksheet.range(
                f'B{current_fill_row}').value = task_project_name
            # 任务ID
            excel_worksheet.range(
                f'C{current_fill_row}').value = task_id
            # 任务标题
            excel_worksheet.range(
                f'D{current_fill_row}').value = task_title
            # 任务链接
            excel_worksheet.range(
                f'E{current_fill_row}').value = task_url
            # 任务状态
            excel_worksheet.range(
                f'F{current_fill_row}').value = "已完成"
            # 优先级
            excel_worksheet.range(
                f'G{current_fill_row}').value = "中"
            # 开始日期
            excel_worksheet.range(
                f'H{current_fill_row}').value = task_detail_data['startDate']
            # 截止日期
            excel_worksheet.range(
                f'I{current_fill_row}').value = task_detail_data['dueDate']
            # 处理人
            excel_worksheet.range(
                f'J{current_fill_row}').value = HANDLER_NAME
            # 完成时间
            excel_worksheet.range(
                f'M{current_fill_row}').value = task_detail_data['dueDate']

            # 预估工时
            try:
                excel_worksheet.range(f'K{current_fill_row}').value = float(
                    est_work_hours_str)
            except (ValueError, TypeError):
                excel_worksheet.range(f'K{current_fill_row}').value = 0.0

            # 完成工时
            try:
                excel_worksheet.range(f'L{current_fill_row}').value = float(
                    task_detail_data['customField1'])
            except (ValueError, TypeError):
                excel_worksheet.range(f'L{current_fill_row}').value = 0.0

            # 自评时长
            try:
                excel_worksheet.range(f'N{current_fill_row}').value = float(
                    task_self_est_hours_str)
            except (ValueError, TypeError):
                excel_worksheet.range(f'N{current_fill_row}').value = 0.0

            print(
                f"序号: {task_serial_num} | 项目: {task_project_name} | 任务: {task_title}")
            current_fill_row += 1
            success_fill_count += 1

        excel_workbook.save()
        return success_fill_count
    except Exception as e:
        print(f"\n❌ Excel数据填充失败：{str(e)}")
        raise
    finally:
        if excel_workbook is not None:
            excel_workbook.close()
        if excel_app is not None:
            excel_app.quit()


def create_excel_from_template(template_file_path):
    """
    基于指定的Excel模板，创建带时间命名的新周报文件
    :param template_file_path: Excel模板文件的完整路径
    :return: 新建Excel文件的完整路径
    """
    excel_app = None
    template_workbook = None
    try:
        # 校验模板文件是否存在
        if not os.path.exists(template_file_path):
            raise FileNotFoundError(f"模板文件不存在，请检查路径：{template_file_path}")

        # 获取当前时间维度信息
        current_datetime = datetime.datetime.now()
        current_year = current_datetime.year
        current_month = current_datetime.month
        current_month_week = get_month_week()

        target_save_dir = os.path.join(
            TARGET_TEMPLATE_DIR_PATH,
            f"{current_year}年",
            f"{current_month}月"
        )
        os.makedirs(target_save_dir, exist_ok=True)

        # 构建新文件名称和完整路径
        new_excel_filename = f"{HANDLER_NAME}-{current_month}月第{current_month_week}周周报.xlsx"
        new_excel_filepath = os.path.join(target_save_dir, new_excel_filename)

        # 打开模板并另存为新文件
        excel_app = xw.App(visible=False, add_book=False)
        template_workbook = excel_app.books.open(template_file_path)
        template_workbook.save(new_excel_filepath)

        return new_excel_filepath
    except Exception as e:
        print(f"\n❌ Excel文件创建失败：{str(e)}")
        raise
    finally:
        # 释放Excel模板资源
        if template_workbook is not None:
            template_workbook.close()
        if excel_app is not None:
            excel_app.quit()


def main():
    try:
        new_excel_path = create_excel_from_template(TEMPLATE_PATH)

        matched_tasks, unrecognized_lines = parse_task_text(
            TEXT, WXWORK_FILL_URL)
        total_recognized_count = len(matched_tasks)

        total_filled_count = fill_task_data_to_excel(
            matched_tasks, new_excel_path)

        print(f"文本中识别到的任务数据条数 ：{total_recognized_count} 条")
        print(f"成功填充到Excel的任务条数 ：{total_filled_count} 条")
        print(f"文本中未识别的无效行数量  ：{len(unrecognized_lines)} 条")

        if unrecognized_lines:
            print("\n未识别的无效数据行：")
            for idx, unrecognized_line in enumerate(unrecognized_lines, 1):
                print(f"   {idx}. {unrecognized_line}")

        print(f"文件存储路径：{new_excel_path}")
    except Exception as e:
        print(f"❌ 程序执行异常终止：{str(e)}")
        raise


if __name__ == "__main__":
    main()
