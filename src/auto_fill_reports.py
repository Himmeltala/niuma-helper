import re
import datetime
import os
import xlwings as xw
from common.cookie_util import get_mission_info
from common.time_util import convert_iso8601_to_hours, get_month_week
from common.validate_util import check_url, check_mission
from config.common import HANDLER_NAME, TEMPLATE_PATH, TARGET_TEMPLATE_DIR_PATH
from config.regex import WXWORK_FILL_URL

TEXT = """
"""


def parse_text(text, pattern):
    """
    解析文本内容，识别匹配和未匹配的行
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
                check_url(line)

    return matches, unrecognized_lines


def fill_excel(matches, excel_path):
    """
    xlwings操作Excel，填充数据
    :param matches: 匹配结果列表
    :param excel_path: 表格文件路径（支持.xlsx/.xls）
    :return: 填充的数据条数
    """
    next_row = 5  # 从第5行开始填充
    app = None
    wb = None
    ws = None
    filled_count = 0

    try:
        # 启动Excel应用（隐藏窗口）
        app = xw.App(visible=False, add_book=False)
        # 打开工作簿（只读模式关闭，允许写入）
        wb = app.books.open(excel_path, read_only=False)
        ws = wb.sheets[0]  # 获取第一个工作表

        # 遍历匹配结果，填充单元格
        for match in matches:
            is_valid = check_mission(match)
            if not is_valid:
                continue

            # 提取匹配组数据
            line_num = match.group(1).strip()
            mission_comptime = match.group(2).strip()
            mission_proname = match.group(3).strip()
            mission_title = match.group(4).strip()
            mission_link = match.group(5).strip()

            # 提取任务ID
            mission_link_mt = re.search(r'/wp/(\d+)', mission_link)
            mission_id = ''
            if mission_link_mt:
                mission_id = mission_link_mt.group(1)
            mission_id = str(int(mission_id)) if mission_id else ''

            if not mission_id:
                continue

            # 获取任务详情
            data = get_mission_info(mission_id)
            html = data['description']['html']

            # 自评时长
            estimated_time = data['estimatedTime']
            estimated_time = convert_iso8601_to_hours(estimated_time)

            # 预估工时/时长
            estimated_duration_match = re.search(
                r'<p class="op-uc-p">预估工时/时长[：:]\s*(\d+\.?\d*).*?</p>',
                html,
                re.S
            )
            estimated_duration = estimated_duration_match.group(
                1).strip() if estimated_duration_match else ''

            # 列B: 项目名称
            ws.range(f'B{next_row}').value = mission_proname
            # 列C: 任务ID
            ws.range(f'C{next_row}').value = mission_id
            # 列D: 任务标题
            ws.range(f'D{next_row}').value = mission_title
            # 列E: 任务链接
            ws.range(f'E{next_row}').value = mission_link
            # 列F: 任务状态
            ws.range(f'F{next_row}').value = "已完成"
            # 列G: 优先级
            ws.range(f'G{next_row}').value = "中"
            # 列H: 开始日期
            ws.range(f'H{next_row}').value = data['startDate']
            # 列I: 截止日期
            ws.range(f'I{next_row}').value = data['dueDate']
            # 列J: 处理人
            ws.range(f'J{next_row}').value = HANDLER_NAME
            # 列K: 预估工时
            try:
                ws.range(f'K{next_row}').value = float(estimated_duration)
            except (ValueError, TypeError):
                ws.range(f'K{next_row}').value = 0
            # 列L: 完成工时
            try:
                ws.range(f'L{next_row}').value = float(data['customField1'])
            except (ValueError, TypeError):
                ws.range(f'L{next_row}').value = 0
            # 列M: 完成时间
            ws.range(f'M{next_row}').value = data['dueDate']
            # 列N: 完成自评时长
            try:
                ws.range(f'N{next_row}').value = float(mission_comptime)
            except (ValueError, TypeError):
                ws.range(f'N{next_row}').value = 0

            print(
                f"序号: {line_num}, 项目名称: {mission_proname}, 任务标题: {mission_title}")
            next_row += 1
            filled_count += 1

        # 保存工作簿
        wb.save()
        return filled_count

    except Exception as e:
        print(f"填充Excel表格失败：{e}")
        raise
    finally:
        # 确保资源释放
        if wb is not None:
            wb.close()
        if app is not None:
            app.quit()


def create_file_from_template(template_path):
    """
    基于Excel模板创建新文件
    :param template_path: 模板文件路径
    :return: 新文件的路径
    """
    app = None
    wb_template = None
    try:
        # 验证模板文件是否存在
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在：{template_path}")

        # 获取当前时间信息
        current_dt = datetime.datetime.now()
        current_month = current_dt.month
        current_year = current_dt.year
        current_week = get_month_week()

        # 构建目标目录并自动创建
        target_dir = os.path.join(
            TARGET_TEMPLATE_DIR_PATH,
            f"{current_year}年",
            f"{current_month}月"
        )
        os.makedirs(target_dir, exist_ok=True)

        # 构建新文件路径
        new_file_name = f"{HANDLER_NAME}-{current_month}月第{current_week}周周报.xlsx"
        new_file_path = os.path.join(target_dir, new_file_name)

        # 打开模板并另存为新文件
        app = xw.App(visible=False, add_book=False)
        wb_template = app.books.open(template_path)
        wb_template.save(new_file_path)

        print(f"Excel新文件创建成功：{new_file_path}")
        return new_file_path

    except Exception as e:
        print(f"创建Excel文件失败：{e}")
        raise
    finally:
        # 释放资源
        if wb_template is not None:
            wb_template.close()
        if app is not None:
            app.quit()


def main():
    try:
        # 基于模板创建新Excel文件
        new_file_path = create_file_from_template(TEMPLATE_PATH)
        # 解析文本数据
        matches, unrecognized_lines = parse_text(TEXT, WXWORK_FILL_URL)
        recognized_count = len(matches)
        # 填充Excel表格
        filled_count = fill_excel(matches, new_file_path)

        print(f"识别的数据条数: {recognized_count}")
        print(f"填充到Excel表格的数据条数: {filled_count}")
        print(f"未识别的数据条数: {len(unrecognized_lines)}")
        if unrecognized_lines:
            print("未识别的行：")
            for line in unrecognized_lines:
                print(f"  - {line}")
    except Exception as e:
        print(f"程序执行失败：{e}")


if __name__ == "__main__":
    main()
