import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import webbrowser
import re
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

def select_pptx():
    print("按钮点击事件触发")
    pptx_path = filedialog.askopenfilename(filetypes=[("PPTX files", "*.pptx")])
    if pptx_path:
        print(f"选择的文件路径: {pptx_path}")
        process_pptx(pptx_path)

def process_pptx(pptx_path):
    pptx_name = os.path.splitext(os.path.basename(pptx_path))[0]

    if save_to_new_directory.get():
        # 把这个内容修改成您的obsidian地址，或者去掉这个
        new_folder = os.path.join(r'D:\notes\Obsidian_data\ppt2md', pptx_name)
    else:
        pptx_dir = os.path.dirname(pptx_path)
        new_folder = os.path.join(pptx_dir, pptx_name)

    images_folder = os.path.join(new_folder, 'images')

    try:
        os.makedirs(new_folder, exist_ok=True)
        print(f"文件夹已创建: {new_folder}")
        convert_pptx_to_md(pptx_path, new_folder, pptx_name, images_folder)
        messagebox.showinfo("成功", f"处理完成！文件夹已创建：{new_folder}")
        open_directory(new_folder)  # 打开目录
    except Exception as e:
        messagebox.showerror("错误", f"发生错误：{e}")
        print(f"错误信息: {e}")

def convert_pptx_to_md(pptx_path, new_folder, pptx_name, images_folder):
    # 构建输出的 markdown 文件路径
    md_file = os.path.join(new_folder, f"{pptx_name}.md")

    # 构建命令行参数列表
    cmd = [
        "pptx2md",
        pptx_path,
        "-o", md_file
    ]



    if disable_escaping.get():
        cmd.append("--disable-escaping")

    if disable_notes.get():
        cmd.append("--disable-notes")

    if disable_wmf.get():
        cmd.append("--disable-wmf")

    if disable_color.get():
        cmd.append("--disable-color")

    if enable_slides.get():
        cmd.append("--enable-slides")

    if wiki.get():
        cmd.append("--wiki")

    if mdk.get():
        cmd.append("--mdk")

    if qmd.get():
        cmd.append("--qmd")

    # 执行命令并捕获输出和错误信息
    try:
        result = subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        print("stdout:", result.stdout)
        print("stderr:", result.stderr)
    except subprocess.CalledProcessError as e:
        print("错误信息:", e.stderr)
        messagebox.showerror("错误", f"pptx2md 转换失败：{e.stderr}")
    else:
        print("转换成功！")
        replace_backslashes(md_file)

def replace_backslashes(md_file):
    try:
        with open(md_file, 'r', encoding='utf-8') as file:
            content = file.read()
        content = content.replace('%5C', '/')
        content = content.replace('', '')  # 替换 "" 为空
        content = content.replace('__', '')  # 替换 "__" 为空

        if disable_image.get():
            content = re.sub(r'!\[.*?\]\([^)]*\)', '', content)  # 替换 ![]() 为空

        with open(md_file, 'w', encoding='utf-8') as file:
            file.write(content)
        print(f"文件中的 '\\' 已成功替换为 '/': {md_file}")
        print(f"文件中的 '' 和 '__' 已成功替换为空: {md_file}")

        if disable_image.get():
            print(f"文件中的 ![]() 已成功替换为空: {md_file}")
    except Exception as e:
        print(f"替换 '\\' 时发生错误: {e}")
        messagebox.showerror("错误", f"替换 '\\' 时发生错误：{e}")

def open_directory(directory_path):
    if os.path.exists(directory_path):
        webbrowser.open(f"file://{directory_path}")
    else:
        messagebox.showerror("错误", f"目录不存在：{directory_path}")

def open_config_window():
    config_window = tk.Toplevel(root)
    config_window.title("配置选项")

    # 添加复选框，用于选择是否禁用特殊字符转义
    checkbox_disable_escaping = tk.Checkbutton(config_window, text="禁用特殊字符转义", variable=disable_escaping)
    checkbox_disable_escaping.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否禁用备注
    checkbox_disable_notes = tk.Checkbutton(config_window, text="禁用备注", variable=disable_notes)
    checkbox_disable_notes.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否禁用wmf格式图片
    checkbox_disable_wmf = tk.Checkbutton(config_window, text="禁用wmf格式图片", variable=disable_wmf)
    checkbox_disable_wmf.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否禁用颜色标签
    checkbox_disable_color = tk.Checkbutton(config_window, text="禁用颜色标签", variable=disable_color)
    checkbox_disable_color.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否启用幻灯片分隔符
    checkbox_enable_slides = tk.Checkbutton(config_window, text="启用幻灯片分隔符", variable=enable_slides)
    checkbox_enable_slides.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否输出md
    checkbox_md = tk.Checkbutton(config_window, text="输出普通md标记语言", variable=md)
    checkbox_md.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否输出tiddlywiki标记语言
    checkbox_wiki = tk.Checkbutton(config_window, text="输出tiddlywiki标记语言", variable=wiki)
    checkbox_wiki.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否输出madoko标记语言
    checkbox_mdk = tk.Checkbutton(config_window, text="输出madoko标记语言", variable=mdk)
    checkbox_mdk.pack(padx=(30, 30), pady=15, anchor=tk.W)

    # 添加复选框，用于选择是否输出qmd标记语言
    checkbox_qmd = tk.Checkbutton(config_window, text="输出qmd标记语言", variable=qmd)
    checkbox_qmd.pack(padx=(30, 30), pady=15, anchor=tk.W)

# 创建主窗口
root = ttk.Window(themename="cosmo")
# root.geometry("800x600")  # 设置窗口大小为 800x600 像素
root.title("PPTX处理工具")

# 定义所有复选框变量
save_to_new_directory = tk.BooleanVar()
disable_image = tk.BooleanVar()
disable_escaping = tk.BooleanVar()
disable_notes = tk.BooleanVar()
disable_wmf = tk.BooleanVar()
disable_color = tk.BooleanVar()
enable_slides = tk.BooleanVar()
md = tk.BooleanVar()
wiki = tk.BooleanVar()
mdk = tk.BooleanVar()
qmd = tk.BooleanVar()

# 添加复选框，用于选择是否将文件保存到新目录
checkbox = tk.Checkbutton(root, text="保存到Obsidian", variable=save_to_new_directory)
checkbox.pack(side=tk.LEFT, padx=(30, 30), pady=15)

# 添加复选框，用于选择是否启用无图模式
checkbox_disable_image = tk.Checkbutton(root, text="无图模式", variable=disable_image)
checkbox_disable_image.pack(side=tk.LEFT, padx=(30, 30), pady=15)

# 创建配置按钮
config_button = tk.Button(root, text="配置", command=open_config_window)
config_button.pack(side=tk.LEFT, padx=(30, 30), pady=15)

# 创建选择文件按钮
select_button = tk.Button(root, text="选择PPTX文件", command=select_pptx)
select_button.pack(side=tk.LEFT, padx=(30, 30), pady=15)

# 运行主窗口
root.mainloop()
