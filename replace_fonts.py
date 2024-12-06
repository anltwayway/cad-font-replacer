# 基于CAD的SCR，LISP，BAT，PS1文件生成器，用于批量替换CAD中的文字样式，采用python开发
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# 支持的 AutoCAD 版本列表
supported_versions = [str(year) for year in range(2013, 2025)]

def generate_scr(style_name, main_font, big_font, height, width_factor, obliquing_angle, reverse, upside_down, vertical, file_name="output.scr"):
    try:
        # 确保文件名以 .scr 结尾
        if not file_name.endswith(".scr"):
            file_name += ".scr"
        
        base_name = os.path.splitext(file_name)[0]
        
        # 生成 SCR 文件
        scr_file_name = f"{base_name}.scr"
        with open(scr_file_name, "w", encoding="gbk") as scr_file:
            scr_file.write(f"-style\n")
            scr_file.write(f"{style_name}\n")  # 文字样式名称
            scr_file.write(f"{main_font},{big_font}\n" if big_font.strip() else f"{main_font}\n")
            scr_file.write(f"{height}\n{width_factor}\n{obliquing_angle}\n{reverse}\n{upside_down}\n")

            # 只有在开启垂直显示设置时写入垂直选项
            if var_enable_vertical.get():
                scr_file.write(f"{vertical}\n")

            scr_file.write("QSAVE\n")

        # 生成 LISP 文件
        lsp_file_name = f"{base_name}.lsp"
        with open(lsp_file_name, "w", encoding="utf-8") as lsp_file:
            lsp_file.write(f"""
(defun ModifyTextStyle ()
  (setq newStyleName "{style_name}") ; 替换为目标样式名称

  ;; 定义需要处理的对象类型
  (setq filter '(
    (0 . "TEXT,MTEXT") ; 选择单行文字、多行文字和标注
  ))

  ;; 获取符合条件的所有对象
  (setq ss (ssget "X" filter))

  ;; 如果找到对象，修改文字样式
  (if ss
    (progn
      (repeat (sslength ss)
        (setq ent (ssname ss 0))
        (setq ed (entget ent))
        (cond
          ;; 单行文字或多行文字
          ((or (eq (cdr (assoc 0 ed)) "TEXT") (eq (cdr (assoc 0 ed)) "MTEXT"))
           (setq ed (subst (cons 7 newStyleName) (assoc 7 ed) ed)) ; 修改样式
          )
        )
        (entmod ed) ; 更新对象
        (ssdel ent ss) ; 从选择集中删除
      )
      (princ (strcat "\\n文字样式已修改为 " newStyleName))
    )
    (princ "\\n未找到符合条件的对象！")
  )
  (princ)
)
            """)

        # 生成 BAT 文件
        bat_file_name = f"{base_name}.bat"
        with open(bat_file_name, "w", encoding="utf-8") as bat_file:
            bat_file.write(f"@echo off\n")
            bat_file.write(f"powershell -NoProfile -ExecutionPolicy Bypass -File \"%~dp0{base_name}.ps1\"\n")
            bat_file.write("pause\n")

        # 获取用户输入的路径
        cad_custom_path = cad_path_var.get().strip()
        if cad_custom_path:
            accoreconsole_path = os.path.join(cad_custom_path, "accoreconsole.exe")
        else:
            # 保留默认版本路径的逻辑
            cad_version = cad_version_var.get()
            accoreconsole_path = f"C:\\Program Files\\Autodesk\\AutoCAD {cad_version}\\accoreconsole.exe"

        # 生成 PS1 文件
        ps1_file_name = f"{base_name}.ps1"
        with open(ps1_file_name, "w", encoding="utf-8") as ps1_file:
            ps1_file.write("# 定义 AutoCAD 的 accoreconsole.exe 路径\n")
            ps1_file.write(f"$accoreconsole = \"{accoreconsole_path}\"\n\n")
            ps1_file.write("# 自动获取当前脚本所在文件夹中的 .scr 文件路径\n")
            ps1_file.write(f"$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath \"{file_name}\"\n\n")
            ps1_file.write("# 获取当前文件夹下的所有 DWG 文件\n")
            ps1_file.write("$dwgFiles = Get-ChildItem -Path $PSScriptRoot -Filter *.dwg\n\n")
            ps1_file.write("# 遍历每个 DWG 文件并运行 accoreconsole\n")
            ps1_file.write("foreach ($file in $dwgFiles) {\n")
            ps1_file.write("    Write-Host \"Processing file: $($file.Name)\"\n")
            ps1_file.write("    & $accoreconsole /i \"$($file.FullName)\" /s \"$scriptPath\"\n")
            ps1_file.write("}\n")

        return scr_file_name, lsp_file_name
    except Exception as e:
        return f"生成文件时出错：{e}"

def on_submit():
    style_name = entry_style_name.get().strip()
    main_font = entry_main_font.get().strip()
    big_font = entry_big_font.get().strip() if var_use_big_font.get() else ""
    height = entry_height.get().strip()
    width_factor = entry_width_factor.get().strip()
    obliquing_angle = entry_obliquing_angle.get().strip()
    # 根据高级选项决定值
    reverse = "Y" if var_advanced_options.get() and var_reverse.get() else "N"
    upside_down = "Y" if var_advanced_options.get() and var_upside_down.get() else "N"
    vertical = "Y" if var_advanced_options.get() and var_vertical.get() else "N"

    file_name = "test"

    # 检查输入是否完整
    if not all([style_name, main_font, height, width_factor, obliquing_angle, file_name]) or (var_use_big_font.get() and not big_font):
        messagebox.showerror("错误", "所有必填字段均需填写完整。")
        return

    try:
        # 调用生成函数
        scr_file_name, lsp_file_name = generate_scr(
            style_name, main_font, big_font, height, width_factor,
            obliquing_angle, reverse, upside_down, vertical, file_name
        )
        # 格式化弹窗消息
        result_message = f"SCR 文件已生成：{scr_file_name}\nLISP 文件已生成：{lsp_file_name}\nbat及ps1文件已生成"
        # messagebox.showinfo("成功", result_message)
    except Exception as e:
        messagebox.showerror("错误", f"生成文件时出错：{e}")

def add_modify_text_commands(file_name, lsp_file_name):
    try:
        if not file_name.endswith(".scr"):
            file_name += ".scr"

        if not os.path.exists(file_name):
            raise FileNotFoundError(f"SCR 文件 {file_name} 不存在，请先生成配置文件！")

        # 替换反斜杠为双反斜杠
        lsp_absolute_path = os.path.abspath(lsp_file_name).replace("\\", "\\\\")

        with open(file_name, "a", encoding="gbk") as scr_file:
            scr_file.write("\n")
            scr_file.write('(setvar "SECURELOAD" 0)\n')  # 临时关闭 SECURELOAD
            scr_file.write(f'(load "{lsp_absolute_path}")\n')
            scr_file.write("(ModifyTextStyle)\n")
            scr_file.write('(setvar "SECURELOAD" 1)\n')  # 恢复默认 SECURELOAD
            scr_file.write("QSAVE\n")  # 保存图纸设置

        return f"已成功追加 LISP 命令到 {file_name}，加载路径: {lsp_absolute_path}"
    except Exception as e:
        return f"追加命令时出错：{e}"

def toggle_vertical_state():
    """根据是否开启垂直显示设置禁用或启用相关选项"""
    state = "normal" if var_enable_vertical.get() else "disabled"
    chk_vertical.config(state=state)

# 创建主窗口
root = tk.Tk()
root.title("SCR 替换字体生成器 v2.0 - 开发者：ZCJ")
root.geometry("800x700")

# 添加变量用于控制是否开启垂直显示选项
var_enable_vertical = tk.BooleanVar(value=False)  # 必须在 Tk 实例创建后定义
var_advanced_options = tk.BooleanVar(value=False)   # 高级选项默认关闭
var_reverse = tk.BooleanVar(value=False)  # 设置为 False，默认为不反向显示
# 添加变量用于控制是否颠倒显示
var_upside_down = tk.BooleanVar(value=False)
# 添加变量用于控制是否垂直显示
var_vertical = tk.BooleanVar(value=False)

# 创建标签和输入框
# 默认值
default_values = {
    "style_name": "txtd",
    "main_font": "fsdb_e'",
    "big_font": "fsdb",
    "height": "0.0",
    "width_factor": "0.7",
    "obliquing_angle": "0.0",
    "file_name": "replace_fonts_test",
}

# 在主窗口中添加下拉框选择 CAD 版本
tk.Label(root, text="选择 AutoCAD 版本").grid(row=14, column=0, padx=10, pady=5, sticky="e")
cad_version_var = tk.StringVar(value="2021")  # 默认设置为 2021
cad_version_dropdown = ttk.Combobox(root, textvariable=cad_version_var, values=supported_versions, state="readonly", width=30)
cad_version_dropdown.grid(row=14, column=1, padx=10, pady=5)

# 添加输入框用于自定义 CAD 路径
tk.Label(root, text="如果 AutoCAD 没有安装在默认C盘目录，请自定义 AutoCAD 路径").grid(row=15, column=0, padx=10, pady=10, sticky="e")
cad_path_var = tk.StringVar(value="")
entry_cad_path = tk.Entry(root, textvariable=cad_path_var, width=40)
entry_cad_path.grid(row=15, column=1, padx=10, pady=5)

# 添加按钮用于选择路径
def browse_cad_path():
    from tkinter import filedialog
    selected_path = filedialog.askdirectory(title="选择 AutoCAD 路径")
    if selected_path:
        cad_path_var.set(selected_path)


def toggle_advanced_options():
    """根据是否开启垂直显示设置禁用或启用相关选项"""
    print(f"高级选项状态：{var_advanced_options.get()}")
    
    if var_advanced_options.get():
        # 显示高级选项控件，并确保它们恢复到原始位置
        lbl_reverse.grid(row=11, column=0, padx=10, pady=5, sticky="e")
        chk_reverse.grid(row=11, column=1)
        lbl_upside_down.grid(row=9, column=0, padx=10, pady=5, sticky="e")
        chk_upside_down.grid(row=9, column=1)
        lbl_vertical.grid(row=10, column=0, padx=10, pady=5, sticky="e")
        chk_vertical.grid(row=10, column=1)

        # 显示高级输入框
        lbl_height.grid(row=7, column=0, padx=10, pady=5, sticky="e")
        entry_height.grid(row=7, column=1, padx=10, pady=5)
        lbl_width_factor.grid(row=5, column=0, padx=10, pady=5, sticky="e")
        entry_width_factor.grid(row=5, column=1, padx=10, pady=5)
        lbl_obliquing_angle.grid(row=6, column=0, padx=10, pady=5, sticky="e")
        entry_obliquing_angle.grid(row=6, column=1, padx=10, pady=5)

        # 显示按钮
        btn_modify_styles.grid(row=16, column=1, columnspan=1, pady=20)
        btn_replace_fonts.grid(row=16, column=2, columnspan=1, pady=20)
        btn_submit.grid(row=16, column=0, columnspan=1, pady=20)
        btn_clear_files.grid(row=17, column=0, columnspan=3, pady=10)

    else:
        # 隐藏高级选项控件
        lbl_reverse.grid_forget()
        chk_reverse.grid_forget()
        lbl_upside_down.grid_forget()
        chk_upside_down.grid_forget()
        lbl_vertical.grid_forget()
        chk_vertical.grid_forget()
        # 隐藏高级输入框
        lbl_height.grid_forget()
        entry_height.grid_forget()
        lbl_width_factor.grid_forget()
        entry_width_factor.grid_forget()
        lbl_obliquing_angle.grid_forget()
        entry_obliquing_angle.grid_forget()
        # 隐藏按钮
        btn_modify_styles.grid_forget()
        btn_replace_fonts.grid_forget()
        btn_submit.grid_forget()
        btn_clear_files.grid_forget()

browse_button = tk.Button(root, text="浏览", command=browse_cad_path)
browse_button.grid(row=15, column=2, padx=5, pady=5)

tk.Label(root, text="要新建或替换的文字样式名称").grid(row=0, column=0, padx=10, pady=5, sticky="e")
style_name_var = tk.StringVar(value=default_values["style_name"])
entry_style_name = tk.Entry(root, textvariable=style_name_var, width=30)
entry_style_name.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="拟使用的主字体文件名称").grid(row=1, column=0, padx=10, pady=5, sticky="e")
main_font_var = tk.StringVar(value=default_values["main_font"])
entry_main_font = tk.Entry(root, textvariable=main_font_var, width=30)
entry_main_font.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="拟使用的大字体文件名称").grid(row=3, column=0, padx=10, pady=5, sticky="e")
big_font_var = tk.StringVar(value=default_values["big_font"])
entry_big_font = tk.Entry(root, textvariable=big_font_var, width=30)
entry_big_font.grid(row=3, column=1, padx=10, pady=5)

# 是否使用大字体复选框
tk.Label(root, text="是否使用大字体").grid(row=2, column=0, padx=10, pady=5, sticky="e")
var_use_big_font = tk.BooleanVar(value=True)
tk.Checkbutton(root, variable=var_use_big_font, command=lambda: entry_big_font.config(state=("normal" if var_use_big_font.get() else "disabled"))).grid(row=2, column=1, padx=10, pady=5, sticky="w")

# 修改大字体输入框的初始状态为禁用
big_font_var = tk.StringVar(value=default_values["big_font"])
entry_big_font = tk.Entry(root, textvariable=big_font_var, width=30, state=("normal" if var_use_big_font.get() else "disabled"))
entry_big_font.grid(row=3, column=1, padx=10, pady=5)

# 文字高度控件
lbl_height = tk.Label(root, text="文字高度")
height_var = tk.StringVar(value=default_values["height"])
entry_height = tk.Entry(root, textvariable=height_var, width=30)

# 文字宽度因子控件
lbl_width_factor = tk.Label(root, text="文字宽度因子")
width_factor_var = tk.StringVar(value=default_values["width_factor"])
entry_width_factor = tk.Entry(root, textvariable=width_factor_var, width=30)

# 倾斜角度控件
lbl_obliquing_angle = tk.Label(root, text="倾斜角度")
obliquing_angle_var = tk.StringVar(value=default_values["obliquing_angle"])
entry_obliquing_angle = tk.Entry(root, textvariable=obliquing_angle_var, width=30)

# 高级选项复选框
tk.Label(root, text="高级选项").grid(row=4, column=0, padx=10, pady=5, sticky="e")
chk_advanced_options = tk.Checkbutton(root, text="显示高级选项", variable=var_advanced_options, command=toggle_advanced_options)
chk_advanced_options.grid(row=4, column=1, padx=10, pady=5, sticky="w")

# 是否反向显示
lbl_reverse = tk.Label(root, text="是否反向显示")
lbl_reverse.grid(row=11, column=0, padx=10, pady=5, sticky="e")
chk_reverse = tk.Checkbutton(root, variable=var_reverse)
chk_reverse.grid(row=11, column=1, padx=10, pady=5, sticky="w")

# 是否颠倒显示
lbl_upside_down = tk.Label(root, text="是否颠倒显示")
lbl_upside_down.grid(row=9, column=0, padx=10, pady=5, sticky="e")
chk_upside_down = tk.Checkbutton(root, variable=var_upside_down)
chk_upside_down.grid(row=9, column=1, padx=10, pady=5, sticky="w")

# 是否垂直显示
lbl_vertical = tk.Label(root, text="是否垂直显示")
lbl_vertical.grid(row=10, column=0, padx=10, pady=5, sticky="e")
chk_vertical = tk.Checkbutton(root, variable=var_vertical)
chk_vertical.grid(row=10, column=1, padx=10, pady=5, sticky="w")

# 提交按钮
btn_submit = tk.Button(root, text="1.生成配置文件", command=on_submit)
btn_submit.grid(row=16, column=0, columnspan=1, pady=20)

def execute_bat_file():
    """
    执行生成的 .bat 文件
    """
    # 获取用户输入的文件名
    file_name = "test"  # 默认文件名
    if not file_name:
        messagebox.showerror("错误", "请先输入文件名并生成配置文件！")
        return

    # 构造基础文件名和 .bat 文件路径
    if not file_name.endswith(".scr"):
        file_name += ".scr"
    base_name = os.path.splitext(file_name)[0]
    bat_file_path = f"{base_name}.bat"

    # 检查文件是否存在并执行
    if os.path.exists(bat_file_path):
        try:
            os.system(f'start "" "{bat_file_path}"')
            messagebox.showinfo("成功", f"成功执行 {bat_file_path}！")
        except Exception as e:
            messagebox.showerror("错误", f"执行 .bat 文件时出错：{e}")
    else:
        messagebox.showerror("错误", f"未找到 .bat 文件：{bat_file_path}\n请先生成配置文件！")

# 在现有代码下方添加新按钮
btn_replace_fonts = tk.Button(root, text="3.点击开始替换", command=execute_bat_file)
btn_replace_fonts.grid(row=16, column=2, columnspan=1, pady=20)

def clear_files():
    """
    清除当前目录下所有 .ps1、.scr 和 .bat 结尾的文件
    """
    file_types = [".ps1", ".scr", ".bat", ".lsp"]
    deleted_files = []

    # 遍历当前目录，找到符合条件的文件并删除
    for file_name in os.listdir():
        if any(file_name.endswith(ext) for ext in file_types):
            try:
                os.remove(file_name)
                deleted_files.append(file_name)
            except Exception as e:
                messagebox.showerror("错误", f"删除文件 {file_name} 时出错：{e}")

    if deleted_files:
        # messagebox.showinfo("成功", f"已成功删除以下文件：\n" + "\n".join(deleted_files))
        messagebox.showinfo("信息", "一键替换成功。")
    else:
        messagebox.showinfo("信息", "未找到任何符合条件的文件。")

btn_clear_files = tk.Button(root, text="清除已生成的配置文件", command=clear_files)
btn_clear_files.grid(row=17, column=0, columnspan=3, pady=10)

def on_modify_styles():
    file_name = "test"  # 获取用户输入的文件名
    style_name = entry_style_name.get().strip()  # 获取文字样式名称

    if not file_name or not style_name:
        messagebox.showerror("错误", "请确保文件名和文字样式名称均已填写！")
        return

    # 检查是否以 .scr 结尾
    if not file_name.endswith(".scr"):
        file_name += ".scr"

    # 检查文件是否存在
    if not os.path.exists(file_name):
        messagebox.showerror("错误", f"SCR 文件 {file_name} 不存在，请先生成该文件！")
        return

    # 调用追加命令函数
    try:
        lsp_file_name = os.path.splitext(file_name)[0] + ".lsp"  # 自动生成对应的 LISP 文件名
        result = add_modify_text_commands(file_name, lsp_file_name)
        # messagebox.showinfo("成功", result)
    except Exception as e:
        messagebox.showerror("错误", f"追加命令时出错：{e}")

btn_modify_styles = tk.Button(root, text="2.CAD文字全改成当前", command=on_modify_styles)
btn_modify_styles.grid(row=16, column=1, columnspan=1, pady=20)

# 逐步执行按钮的功能
def execute_steps():
    # 执行步骤1
    on_submit()
    # 执行步骤2
    on_modify_styles()
    # 执行步骤3
    execute_bat_file()
    # 执行步骤4
    clear_files()

# 新增的逐步执行按钮
btn_execute_steps = tk.Button(root, text="一键替换", command=execute_steps)
btn_execute_steps.grid(row=18, column=0, columnspan=3, pady=20)

# 使用说明

# 创建菜单栏
menu_bar = tk.Menu(root)

# 创建“帮助”菜单
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="使用说明", command=lambda: show_help())
help_menu.add_separator()
help_menu.add_command(label="关于", command=lambda: show_about())

# 将“帮助”菜单添加到菜单栏
menu_bar.add_cascade(label="帮助", menu=help_menu)

# 配置菜单栏
root.config(menu=menu_bar)

# 定义帮助弹窗函数
def show_help():
    help_text = """
    使用说明：
    将此exe文件与要修改的dwg文件放在同一目录下，运行exe文件，支持修改同目录下所有CAD。 
    1. 填写文字样式名称和主字体文件名称。
    2. 根据需要勾选是否使用大字体，并填写大字体名称。
    3. 填写文字高度、宽度因子和倾斜角度。
    4. 勾选是否反向显示或颠倒显示。
    5. 生成 SCR 文件后，可进行图纸文字样式修改。
    6. 清理生成的文件时，点击“清除已生成的配置文件”按钮。
    7. 选项“2.CAD文字全改成当前”不是必须的，此功能可将dwg所有文字对象的文字样式修改为当前设置的文字样式。如果不需要此功能，只需要执行1和3即可。
    8. 如果执行1和3，就只是替换或者新建文字样式，不会修改CAD里面的那些字。
    9. 如果全部执行1 2 3 会替换活新建文字样式，并且把CAD里面的文字都应用当前文字样式。
    """
    messagebox.showinfo("使用说明", help_text)

# 定义关于弹窗函数
def show_about():
    about_text = "SCR 替换字体生成器 v1.1\n开发者：ZCJ\n2024 年"
    messagebox.showinfo("关于", about_text)

# 默认隐藏高级选项控件
lbl_reverse.grid_remove()
chk_reverse.grid_remove()
lbl_upside_down.grid_remove()
chk_upside_down.grid_remove()
lbl_vertical.grid_remove()
chk_vertical.grid_remove()

# 隐藏按钮
btn_modify_styles.grid_remove()
btn_replace_fonts.grid_remove()
btn_submit.grid_remove()
btn_clear_files.grid_remove()

# 运行主循环
root.mainloop()
