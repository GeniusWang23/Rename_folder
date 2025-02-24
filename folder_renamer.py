import os
import pandas as pd
import logging
from tkinter import Tk, Button, Label, StringVar
from tkinter.filedialog import askopenfilename, askdirectory

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


def load_mapping(excel_file_path, sheet_name=0, column_index=1):
    """
    读取Excel文件并构建文件编号与文件名称的双向映射关系。
    假设文件编号在奇数行，文件名称在偶数行。
    """
    try:
        # 读取Excel文件，不使用表头
        df = pd.read_excel(excel_file_path, header=None)
    except Exception as e:
        logging.error(f"读取Excel文件时出错: {e}")
        return None, None

    mapping = {}
    reverse_mapping = {}
    # 遍历Excel的每一行，步长为2
    for i in range(0, len(df), 2):
        if i + 1 < len(df):
            file_id = str(df.iloc[i, column_index]).strip()
            file_name = str(df.iloc[i + 1, column_index]).strip()
            if file_id and file_name:
                mapping[file_id] = file_name
                reverse_mapping[file_name] = file_id
                logging.debug(f"映射添加: {file_id} <-> {file_name}")
            else:
                logging.warning(f"跳过第 {i} 行和第 {i + 1} 行，因为文件编号或文件名称为空。")
        else:
            logging.warning(f"跳过第 {i} 行，因为缺少新文件夹名称。")

    return mapping, reverse_mapping


def rename_folders(target_path, mapping, reverse_mapping):
    """
    遍历target_path下的所有子文件夹，并对每个子文件夹下的文件夹根据映射关系进行重命名。
    """
    success_count = 0
    fail_count = 0
    total_count = 0
    # 遍历当前目录下的所有子文件夹
    for sub_folder in os.listdir(target_path):
        sub_folder_path = os.path.join(target_path, sub_folder)
        if os.path.isdir(sub_folder_path):
            # 遍历子文件夹下的所有文件夹
            for folder in os.listdir(sub_folder_path):
                folder_path = os.path.join(sub_folder_path, folder)
                if os.path.isdir(folder_path):
                    total_count += 1
                    new_name = None
                    if folder in mapping:
                        new_name = mapping[folder]
                    elif folder in reverse_mapping:
                        new_name = reverse_mapping[folder]

                    if new_name:
                        if new_name == folder:
                            # 如果文件夹已经是重命名后的名称，跳过
                            logging.info(f"文件夹 '{folder}' 已经是重命名后的名称，跳过。")
                            continue
                        new_path = os.path.join(sub_folder_path, new_name)
                        # 检查新名称是否已存在
                        if os.path.exists(new_path):
                            logging.error(f"重命名失败，目标文件夹 '{new_name}' 已存在。")
                            fail_count += 1
                            continue
                        try:
                            os.rename(folder_path, new_path)
                            logging.info(f"已重命名: {folder} -> {new_name}")
                            success_count += 1
                        except Exception as e:
                            logging.error(f"重命名 '{folder}' 时出错: {e}")
                            fail_count += 1
                    else:
                        logging.warning(f"未找到 '{folder}' 的映射，跳过。")
                        fail_count += 1

    return success_count, fail_count, total_count


def select_excel_file():
    global excel_file_path
    excel_file_path = askopenfilename(title="选择 Excel 文件", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if excel_file_path:
        excel_path_label.config(text=f"选择的 Excel 文件路径: {excel_file_path}")


def select_target_folder():
    global target_path
    target_path = askdirectory(title="选择要重命名文件夹的上一层文件夹")
    if target_path:
        target_folder_label.config(text=f"选择的文件夹路径: {target_path}")


def start_renaming():
    if not excel_file_path or not target_path:
        result_label.config(text="请先选择 Excel 文件和目标文件夹！")
        return
    # 加载映射关系
    mapping, reverse_mapping = load_mapping(excel_file_path)
    if not mapping or not reverse_mapping:
        result_label.config(text="错误: 未加载到任何有效的映射关系。")
        return
    # 执行重命名
    success_count, fail_count, total_count = rename_folders(target_path, mapping, reverse_mapping)
    not_renamed_count = total_count - success_count
    result_text = f"重命名操作完成。总共有 {total_count} 个文件夹，已重命名数量: {success_count}，未重命名数量: {not_renamed_count}"
    result_label.config(text=result_text)


# 全局变量
excel_file_path = ""
target_path = ""

# 创建主窗口
root = Tk()
root.title("文件夹重命名工具")

# 选择 Excel 文件按钮
excel_button = Button(root, text="选择 Excel 文件", command=select_excel_file)
excel_button.pack(pady=10)

# 显示 Excel 文件路径的标签
excel_path_label = Label(root, text="未选择 Excel 文件")
excel_path_label.pack(pady=5)

# 选择目标文件夹按钮
target_folder_button = Button(root, text="选择目标文件夹", command=select_target_folder)
target_folder_button.pack(pady=10)

# 显示目标文件夹路径的标签
target_folder_label = Label(root, text="未选择目标文件夹")
target_folder_label.pack(pady=5)

# 开始重命名按钮
start_button = Button(root, text="开始重命名", command=start_renaming)
start_button.pack(pady=20)

# 显示重命名结果的标签
result_label = Label(root, text="")
result_label.pack(pady=10)

# 运行主循环
root.mainloop()
