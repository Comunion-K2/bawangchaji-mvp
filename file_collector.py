import os
import shutil
import pandas as pd
from tkinter import messagebox

# 配置路径
source_root = "E:/第十三批整合/"
output_root = os.path.join(source_root, "收集结果")
excel_filename = "target_files.xlsx"

# 创建输出目录
os.makedirs(output_root, exist_ok=True)

# 读取Excel中的编号
try:
    df = pd.read_excel(excel_filename)
except Exception as e:
    messagebox.showerror("错误", f"无法读取 Excel 文件：{e}")
    exit()

keywords = df.iloc[:, 0].astype(str).tolist()
not_found = []

# 遍历编号，查找匹配文件夹
for keyword in keywords:
    found = False
    for root, dirs, files in os.walk(source_root):
        for dir_name in dirs:
            if keyword in dir_name:
                source_path = os.path.join(root, dir_name)
                target_path = os.path.join(output_root, dir_name)
                try:
                    if not os.path.exists(target_path):
                        shutil.copytree(source_path, target_path)
                    found = True
                    break
                except Exception as e:
                    messagebox.showerror("复制失败", f"复制 {source_path} 失败：{e}")
                    break
        if found:
            break
    if not found:
        not_found.append(keyword)

# 输出未找到编号
if not_found:
    pd.DataFrame(not_found, columns=["未找到编号"]).to_excel("未找到编号.xlsx", index=False)
    messagebox.showinfo("完成", f"任务完成！未找到 {len(not_found)} 个编号，已输出为 Excel。")
else:
    messagebox.showinfo("完成", "所有编号对应文件夹均已找到并复制成功！")
