import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import NamedStyle
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import ttk
import os
import random
import sys
import subprocess

# 初始内存中的黑名单数据
memory_blacklist = pd.DataFrame(columns=['订单号', '买家id', '收货人名称', '收货地址', '联系电话', '手机'])

def restart_program():
    """重启当前程序"""
    print("重启程序...")
    subprocess.Popen([sys.executable] + sys.argv)
    sys.exit()


def process_files(file_a, file_b):
    try:
        df1 = pd.read_excel(file_a, dtype=str)  # 读取为字符串以避免科学计数法
        df2 = pd.read_excel(file_b, dtype=str)  # 读取为字符串以避免科学计数法
    except Exception as e:
        messagebox.showerror("错误", f"读取文件失败: {e}")
        return

    # 填充买家名单中的缺失数据
    for column in ['买家id', '收货人名称', '收货地址', '联系电话', '手机']:
        if column in df2.columns:
            df2[column].fillna(
                df2[column].apply(lambda x: str(random.randint(10000, 99999)).zfill(5) if pd.isnull(x) else x),
                inplace=True)

    # 处理订单号的缺失
    if '订单号' in df2.columns:
        df2['订单号'].fillna(
            df2['订单号'].apply(lambda x: str(random.randint(10000, 99999)).zfill(5) if pd.isnull(x) else x),
            inplace=True)

    df2 = df2.dropna(subset=['订单号', '买家id', '收货人名称', '收货地址'], how='all')

    combined_data = {}
    matched_records = set()

    for index, row in df1.iterrows():
        buyer_id = row['买家id']
        recipient_name = row['收货人名称']
        delivery_address = row['收货地址']
        mobile_number = row['手机']
        order_number = row['订单号']

        if (buyer_id, recipient_name, delivery_address, mobile_number, order_number) in matched_records:
            continue

        match = df2[
            (df2['买家id'] == buyer_id) |
            (df2['收货人名称'] == recipient_name) |
            (df2['收货地址'] == delivery_address) |
            (df2['手机'] == mobile_number) |
            (df2['订单号'] == order_number)
            ]

        if not match.empty:
            for _, match_row in match.iterrows():
                matched_order_number = match_row['订单号'] if pd.notna(match_row['订单号']) else '无'
                matched_buyer_id = match_row['买家id'] if pd.notna(match_row['买家id']) else '无'
                matched_recipient_name = match_row['收货人名称'] if pd.notna(match_row['收货人名称']) else '无'
                matched_delivery_address = match_row['收货地址'] if pd.notna(match_row['收货地址']) else '无'
                matched_phone_number = match_row['联系电话'] if pd.notna(match_row['联系电话']) else '无'
                matched_mobile_number = match_row['手机'] if pd.notna(match_row['手机']) else '无'

                # 生成备注信息
                remarks = []
                if order_number == matched_order_number:
                    remarks.append('订单号')
                if buyer_id == matched_buyer_id:
                    remarks.append('买家ID')
                if recipient_name == matched_recipient_name:
                    remarks.append('收货人名称')
                if delivery_address == matched_delivery_address:
                    remarks.append('收货地址')
                if mobile_number == matched_mobile_number:
                    remarks.append('手机')

                remarks_text = ', '.join(remarks) if remarks else '无'

                combined_data[matched_order_number] = [
                    matched_buyer_id,
                    matched_recipient_name,
                    matched_delivery_address,
                    matched_phone_number,
                    matched_mobile_number,
                    remarks_text
                ]
                matched_records.add((buyer_id, recipient_name, delivery_address, mobile_number, order_number))

    update_table(combined_data)


def is_random_value(value):
    return len(value) == 5 and value.isdigit()

def update_table(combined_data):
    for row in tree.get_children():
        tree.delete(row)

    if not combined_data:
        messagebox.showinfo("信息", "没有找到匹配的记录。")

    for order_number, details in combined_data.items():
        # 替换随机生成的五位数为“无”
        row_data = [
            order_number if not is_random_value(str(order_number)) else '无'
        ] + [
            value if not is_random_value(str(value)) else '无'
            for value in details
        ]
        tree.insert('', 'end', values=row_data)




def on_drop_a(event):
    global file_a_path
    new_file_a = root.tk.splitlist(event.data)[0]
    file_a_label.config(text=f"黑名单文件已拖入: {new_file_a}")
    file_a_label.config(bg="lightgreen")

    if os.path.isfile(file_a_path):
        merge_and_save(file_a_path, new_file_a)
        file_a_path = new_file_a
        file_a_label.config(text=f"黑名单文件: {file_a_path}")
        messagebox.showinfo("成功", f"黑名单文件已更新: {file_a_path}")
    else:
        file_a_path = new_file_a
        file_a_label.config(text=f"黑名单文件: {file_a_path}")
        messagebox.showinfo("成功", f"黑名单文件已设置: {file_a_path}")
    restart_program()

def on_drop_b(event):
    global file_b_path
    file_b = root.tk.splitlist(event.data)[0]
    file_b_label.config(text=f"买家名单文件已拖入: {file_b}")
    file_b_label.config(bg="lightgreen")
    file_b_path = file_b

    if file_a_path and os.path.isfile(file_a_path):
        process_files(file_a_path, file_b_path)
    else:
        messagebox.showwarning("警告", "请先拖入有效的黑名单文件。")

def merge_and_save(old_file, new_file):
    global memory_blacklist

    if not os.path.isfile(old_file):
        messagebox.showerror("错误", f"旧文件不存在: {old_file}")
        return

    old_df = pd.read_excel(old_file, dtype=str)  # 读取为字符串以避免科学计数法
    new_df = pd.read_excel(new_file, dtype=str)  # 读取为字符串以避免科学计数法

    combined_df = pd.concat([old_df, new_df, memory_blacklist]).drop_duplicates().reset_index(drop=True)

    with pd.ExcelWriter(old_file, engine='openpyxl') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']

        text_style = NamedStyle(name='text_style', number_format='@')

        for col in combined_df.columns:
            col_idx = combined_df.columns.get_loc(col) + 1
            for row in range(2, len(combined_df) + 2):
                cell = worksheet.cell(row=row, column=col_idx)
                if isinstance(cell.value, str) and cell.value.isdigit():
                    cell.style = text_style  # 设置单元格样式为文本

    memory_blacklist = combined_df

    print("合并后的黑名单数据:")
    print(memory_blacklist)

def add_blacklist_entry():
    global memory_blacklist

    order_number = order_number_entry.get().strip()
    buyer_id = buyer_id_entry.get().strip()
    recipient_name = recipient_name_entry.get().strip()
    delivery_address = delivery_address_entry.get().strip()
    phone_number = phone_number_entry.get().strip()
    mobile_number = mobile_number_entry.get().strip()

    if not order_number:
        order_number = str(random.randint(10000, 99999)).zfill(5)  # 生成随机订单号

    new_entry = {
        '订单号': order_number,
        '买家id': buyer_id if buyer_id else "",
        '收货人名称': recipient_name if recipient_name else "",
        '收货地址': delivery_address if delivery_address else "",
        '联系电话': phone_number if phone_number else "",
        '手机': mobile_number if mobile_number else ""
    }

    new_entry_df = pd.DataFrame([new_entry])

    excel_file = file_a_path

    try:
        df_existing = pd.read_excel(excel_file, dtype=str)  # 读取为字符串以避免科学计数法
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=new_entry.keys())  # 如果文件不存在，则创建一个新的 DataFrame

    df_new = pd.DataFrame([new_entry])

    df_combined = pd.concat([df_existing, df_new], ignore_index=True).drop_duplicates().reset_index(drop=True)

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
        df_combined.to_excel(writer, index=False, sheet_name='Sheet1')

    messagebox.showinfo("成功", "条目已添加到黑名单。")

    order_number_entry.delete(0, tk.END)
    buyer_id_entry.delete(0, tk.END)
    recipient_name_entry.delete(0, tk.END)
    delivery_address_entry.delete(0, tk.END)
    phone_number_entry.delete(0, tk.END)
    mobile_number_entry.delete(0, tk.END)
def search_blacklist():
    global file_a_path

    # 获取输入的查询条件
    query_order_number = order_number_search_entry.get().strip()
    query_buyer_id = buyer_id_search_entry.get().strip()
    query_recipient_name = recipient_name_search_entry.get().strip()
    query_delivery_address = delivery_address_search_entry.get().strip()
    query_phone_number = phone_number_search_entry.get().strip()
    query_mobile_number = mobile_number_search_entry.get().strip()

    # 打印查询条件用于调试
    print(f"查询条件 - 订单号: {query_order_number}, 买家ID: {query_buyer_id}, 收货人名称: {query_recipient_name}, 收货地址: {query_delivery_address}, 联系电话: {query_phone_number}, 手机: {query_mobile_number}")

    # 读取黑名单数据
    try:
        df_blacklist = pd.read_excel(file_a_path, dtype=str)  # 读取为字符串以避免科学计数法
    except Exception as e:
        messagebox.showerror("错误", f"读取黑名单文件失败: {e}")
        return

    # 创建查询条件
    conditions = [
        df_blacklist['订单号'].str.contains(query_order_number, na=False, case=False) if query_order_number else True,
        df_blacklist['买家id'].str.contains(query_buyer_id, na=False, case=False) if query_buyer_id else True,
        df_blacklist['收货人名称'].str.contains(query_recipient_name, na=False, case=False) if query_recipient_name else True,
        df_blacklist['收货地址'].str.contains(query_delivery_address, na=False, case=False) if query_delivery_address else True,
        df_blacklist['联系电话'].str.contains(query_phone_number, na=False, case=False) if query_phone_number else True,
        df_blacklist['手机'].str.contains(query_mobile_number, na=False, case=False) if query_mobile_number else True
    ]

    # 将所有条件组合起来
    combined_condition = conditions[0]
    for condition in conditions[1:]:
        combined_condition &= condition

    # 根据组合条件筛选结果
    search_result = df_blacklist[combined_condition]

    # 替换 NaN 为 '无'
    search_result = search_result.fillna('无')

    # 更新表格显示
    for row in tree.get_children():
        tree.delete(row)

    if search_result.empty:
        messagebox.showinfo("信息", "没有找到匹配的记录。")
    else:
        for _, row in search_result.iterrows():
            row_data = [
                row['订单号'] if pd.notna(row['订单号']) else '无',
                row['买家id'] if pd.notna(row['买家id']) else '无',
                row['收货人名称'] if pd.notna(row['收货人名称']) else '无',
                row['收货地址'] if pd.notna(row['收货地址']) else '无',
                row['联系电话'] if pd.notna(row['联系电话']) else '无',
                row['手机'] if pd.notna(row['手机']) else '无'
            ]
            tree.insert('', 'end', values=row_data)



root = TkinterDnD.Tk()
root.title("数据匹配工具")
root.geometry("1400x900")

default_file_a = 'a.xlsx'
file_a_path = os.path.join(os.getcwd(), default_file_a)

main_frame = tk.Frame(root)
main_frame.pack(expand=True, fill='both', pady=10, padx=10)

left_frame = tk.Frame(main_frame)
left_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

right_frame = tk.Frame(main_frame)
right_frame.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')

file_a_label = tk.Label(left_frame, text=f"黑名单文件: {file_a_path}", fg="blue", bg="lightgrey")
file_a_label.grid(row=1, column=0, pady=5)

file_b_label = tk.Label(left_frame, text="买家名单文件: 未选择", fg="blue", bg="lightgrey")
file_b_label.grid(row=1, column=1, pady=5)

drop_area_a = tk.Label(left_frame, text=f"更新黑名单文件区域\n当前黑名单: {file_a_path}", bg="lightgrey", width=40,
                       height=15, relief="raised")
drop_area_a.grid(row=0, column=0, pady=10)
drop_area_a.drop_target_register(DND_FILES)
drop_area_a.dnd_bind('<<Drop>>', on_drop_a)

drop_area_b = tk.Label(left_frame, text="买家名单文件区域", bg="lightgrey", width=40, height=15, relief="raised")
drop_area_b.grid(row=0, column=1, pady=10)
drop_area_b.drop_target_register(DND_FILES)
drop_area_b.dnd_bind('<<Drop>>', on_drop_b)

# 添加和查询板块位于同一排
add_frame = tk.Frame(right_frame)
add_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

tk.Label(add_frame, text="订单号:").grid(row=0, column=0, padx=5, pady=5)
order_number_entry = tk.Entry(add_frame)
order_number_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(add_frame, text="买家ID:").grid(row=1, column=0, padx=5, pady=5)
buyer_id_entry = tk.Entry(add_frame)
buyer_id_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(add_frame, text="收货人名称:").grid(row=2, column=0, padx=5, pady=5)
recipient_name_entry = tk.Entry(add_frame)
recipient_name_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(add_frame, text="收货地址:").grid(row=3, column=0, padx=5, pady=5)
delivery_address_entry = tk.Entry(add_frame)
delivery_address_entry.grid(row=3, column=1, padx=5, pady=5)

tk.Label(add_frame, text="联系电话:").grid(row=4, column=0, padx=5, pady=5)
phone_number_entry = tk.Entry(add_frame)
phone_number_entry.grid(row=4, column=1, padx=5, pady=5)

tk.Label(add_frame, text="手机:").grid(row=5, column=0, padx=5, pady=5)
mobile_number_entry = tk.Entry(add_frame)
mobile_number_entry.grid(row=5, column=1, padx=5, pady=5)

add_button = tk.Button(add_frame, text="添加到黑名单", command=add_blacklist_entry)
add_button.grid(row=6, column=0, columnspan=2, pady=10)

# 搜索板块
search_frame = tk.Frame(right_frame)
search_frame.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')

tk.Label(search_frame, text="订单号:").grid(row=0, column=0, padx=5, pady=5)
order_number_search_entry = tk.Entry(search_frame)
order_number_search_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(search_frame, text="买家ID:").grid(row=1, column=0, padx=5, pady=5)
buyer_id_search_entry = tk.Entry(search_frame)
buyer_id_search_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(search_frame, text="收货人名称:").grid(row=2, column=0, padx=5, pady=5)
recipient_name_search_entry = tk.Entry(search_frame)
recipient_name_search_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(search_frame, text="收货地址:").grid(row=3, column=0, padx=5, pady=5)
delivery_address_search_entry = tk.Entry(search_frame)
delivery_address_search_entry.grid(row=3, column=1, padx=5, pady=5)

tk.Label(search_frame, text="联系电话:").grid(row=4, column=0, padx=5, pady=5)
phone_number_search_entry = tk.Entry(search_frame)
phone_number_search_entry.grid(row=4, column=1, padx=5, pady=5)

tk.Label(search_frame, text="手机:").grid(row=5, column=0, padx=5, pady=5)
mobile_number_search_entry = tk.Entry(search_frame)
mobile_number_search_entry.grid(row=5, column=1, padx=5, pady=5)

search_button = tk.Button(search_frame, text="查询黑名单", command=search_blacklist)
search_button.grid(row=6, column=0, columnspan=2, pady=10)

columns = ['订单号', '买家id', '收货人名称', '收货地址', '联系电话', '手机', '重复项']
table_frame = tk.Frame(root)
table_frame.pack(expand=True, fill='both', pady=10, padx=10)

tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=200, anchor='w')
tree.pack(side='left', fill='both', expand=True)


table_frame.bind('<Configure>', lambda e: tree.column('#0', width=table_frame.winfo_width() // len(columns)))


root.mainloop()
