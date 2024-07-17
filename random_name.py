
from tkinter import Tk , Button, Label, filedialog,BOTH,BOTTOM,TOP
from pandas import read_json, read_excel
import os.path
# 缓存文件路径
cache_file = "students_cache.json"

# 检查缓存文件是否存在，并加载数据
def load_from_cache():
    if os.path.exists(cache_file):
        with open(cache_file, 'r') as file:
            global students
            students = read_json(file.read(), orient='records')
        display_message("从缓存加载数据成功！")
    else:
        display_message("缓存文件不存在，需要加载新的Excel文件。")

# 将数据写入缓存文件
def save_to_cache():
    with open(cache_file, 'w') as file:
        json_str = students.to_json(orient='records')
        file.write(json_str)
    display_message("数据已缓存！")

# 随机选择学生
def random_select():
    global i
    if not students.empty:
        # 如果之前已经点过名，确保已点名的学生列表不为空
        if selected_students:
            # 从学生数据中排除已点过名的学生
            available_students = students[~students[0].isin(selected_students)]
        else:
            # 如果没有点过名，使用全部学生数据
            available_students = students

        # 如果还有未点过名的学生
        if not available_students.empty:
            selected_student = available_students.sample(1)
            # 记录这次点名的学生
            selected_students.append(selected_student.iloc[0, 0])
            i=i+1
            display_message(f"第{i}人，姓名：{selected_student.iloc[0, 1]}，学号：{selected_student.iloc[0, 0]}")

        else:
            # 如果所有学生都被点过名了
            display_message("所有学生都已点过名。")
    else:
        display_message("没有学生数据，请先导入Excel文件。")

# 在程序开始时初始化已点名的学生列表
selected_students = []

# 导入Excel文件
def load_excel():
    global students
    # 弹出文件选择对话框
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        students = read_excel(file_path, header=None)  # 假设没有表头
        save_to_cache()  # 更新缓存
        display_message("文件加载成功！")

# 显示消息
def display_message(message):
    message_label.config(text=message, font=('Helvetica',36))

# 创建主窗口

root = Tk()
root.title("随机点名程序")
root.geometry("1000x300+300+300")  # 设置窗口大小和初始位置
# 创建按钮
load_button = Button(root, text="导入数据", command=load_excel)
load_button.pack(side=TOP, padx=0, pady=0)

# 创建消息显示标签
message_label = Label(root, text="", font=('Helvetica', 36), bg='black', fg='white')
message_label.pack(expand=True, fill=BOTH)

select_button = Button(root, text="随机点名", font=('Helvetica', 20) ,command=random_select)
select_button.pack(side=BOTTOM, pady=20)
i=0
# 首次运行时加载缓存数据
load_from_cache()

# 运行主循环
root.mainloop()