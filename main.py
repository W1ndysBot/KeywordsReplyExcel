# script/KeywordsReplyExcel/main.py

import logging
import os
import sys
import xlrd

# 添加项目根目录到sys.path
sys.path.append(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
)

from app.config import *
from app.api import *
from app.switch import load_switch, save_switch


# 数据存储路径，实际开发时，请将KeywordsReplyExcel替换为具体的数据存放路径
DATA_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    "data",
    "KeywordsReplyExcel",
)


# 查看功能开关状态
def load_function_status(group_id):
    return load_switch(group_id, "KeywordsReplyExcel")


# 保存功能开关状态
def save_function_status(group_id, status):
    save_switch(group_id, "KeywordsReplyExcel", status)


# 处理元事件，用于启动时确保数据目录存在
async def handle_KeywordsReplyExcel_meta_event(websocket, msg):
    os.makedirs(DATA_DIR, exist_ok=True)


# 在Excel寻找符合的关键词回复数据
def get_first_and_second_column_values_by_keyword(keyword):
    # 获取第一张表
    df = xlrd.open_workbook(os.path.join(DATA_DIR, "data.xls"))
    sheet = df.sheet_by_index(0)
    # 初始化结果列表
    result = []
    # 遍历所有行
    for row_idx in range(sheet.nrows):
        # 获取第一行的值
        first_col_value = sheet.cell_value(row_idx, 0)
        # 检查第一列是否包含关键字
        if keyword in str(first_col_value):
            # 如果包含，获取第二列的值
            second_col_value = sheet.cell_value(row_idx, 1)
            # 将第一列和第二列的值作为元组添加到结果列表
            result.append((first_col_value, second_col_value))
    return result


# 处理开关状态
async def toggle_function_status(websocket, group_id, message_id, authorized):
    if not authorized:
        await send_group_msg(
            websocket,
            group_id,
            f"[CQ:reply,id={message_id}]❌❌❌你没有权限对KeywordsReplyExcel功能进行操作,请联系管理员。",
        )
        return

    if load_function_status(group_id):
        save_function_status(group_id, False)
        await send_group_msg(
            websocket,
            group_id,
            f"[CQ:reply,id={message_id}]🚫🚫🚫KeywordsReplyExcel功能已关闭",
        )
    else:
        save_function_status(group_id, True)
        await send_group_msg(
            websocket,
            group_id,
            f"[CQ:reply,id={message_id}]✅✅✅KeywordsReplyExcel功能已开启",
        )


# 群消息处理函数
async def handle_KeywordsReplyExcel_group_message(websocket, msg):
    # 确保数据目录存在
    os.makedirs(DATA_DIR, exist_ok=True)
    try:
        user_id = str(msg.get("user_id"))
        group_id = str(msg.get("group_id"))
        raw_message = str(msg.get("raw_message"))
        role = str(msg.get("sender", {}).get("role"))
        message_id = str(msg.get("message_id"))
        authorized = user_id in owner_id

        # 是否是开启命令
        if raw_message.startswith("kre"):
            await toggle_function_status(websocket, group_id, message_id, authorized)
        else:
            message = f"[CQ:reply,id={message_id}]"
            # 获取关键词回复数据
            results = get_first_and_second_column_values_by_keyword(raw_message)
            if results:
                for item in results:
                    message += f"文件名: {item[0]} -> 链接: {item[1]}\n"
                await send_group_msg(
                    websocket,
                    group_id,
                    message,
                )
    except Exception as e:
        logging.error(f"处理KeywordsReplyExcel群消息失败: {e}")
        return
