import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from pptx import Presentation
import copy
import os
from datetime import datetime

# ---------------- PPT 操作函数 ----------------
def duplicate_slide(prs, slide):
    blank_slide_layout = prs.slide_layouts[6]
    new_slide = prs.slides.add_slide(blank_slide_layout)
    for shape in slide.shapes:
        el = shape.element
        new_slide.shapes._spTree.insert_element_before(copy.deepcopy(el), 'p:extLst')
    return new_slide

def merge_runs_in_paragraph(para):
    runs = para.runs
    if len(runs) <= 1:
        return
    full_text = "".join(run.text for run in runs)
    runs[0].text = full_text
    for run in runs[1:]:
        run._r.getparent().remove(run._r)

def replace_text_in_shape(shape, replacements):
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        merge_runs_in_paragraph(para)
        for run in para.runs:
            for placeholder, value in replacements.items():
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(value))

def generate_ppt(excel_path, ppt_path, output_folder):
    df = pd.read_excel(excel_path)
    prs = Presentation(ppt_path)
    template_slide = prs.slides[0]

    for index, row in df.iterrows():
        slide = duplicate_slide(prs, template_slide)

        # 自动匹配 Excel 列名到占位符 [列名]
        replacements = {}
        for col in df.columns:
            placeholder = f"[{col}]"
            replacements[placeholder] = str(row[col])

        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

    # 删除模板页
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

    # 输出文件名
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    template_name = os.path.splitext(os.path.basename(ppt_path))[0]
    new_filename = f"{template_name}_生成_{timestamp}.pptx"
    new_path = os.path.join(output_folder, new_filename)
    prs.save(new_path)
    return new_path

# ---------------- Tkinter GUI ----------------
def select_excel():
    path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(0, path)

def select_ppt():
    path = filedialog.askopenfilename(filetypes=[("PPT Files","*.pptx")])
    ppt_entry.delete(0, tk.END)
    ppt_entry.insert(0, path)

def select_output():
    path = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, path)

def generate():
    excel_path = excel_entry.get()
    ppt_path = ppt_entry.get()
    output_folder = output_entry.get()
    if not excel_path or not ppt_path or not output_folder:
        messagebox.showwarning("提示", "请填写 Excel 文件、PPT 模板和输出文件夹")
        return
    try:
        new_path = generate_ppt(excel_path, ppt_path, output_folder)
        messagebox.showinfo("完成", f"PPT生成完成！\n文件路径：{new_path}")
    except Exception as e:
        messagebox.showerror("错误", str(e))

root = tk.Tk()
root.title("PPT生成器懒人专用")

tk.Label(root, text="Excel 文件:").grid(row=0,column=0,padx=5,pady=5, sticky='e')
excel_entry = tk.Entry(root, width=50); excel_entry.grid(row=0,column=1)
tk.Button(root,text="选择", command=select_excel).grid(row=0,column=2,padx=5,pady=5)

tk.Label(root, text="PPT 模板:").grid(row=1,column=0,padx=5,pady=5, sticky='e')
ppt_entry = tk.Entry(root, width=50); ppt_entry.grid(row=1,column=1)
tk.Button(root,text="选择", command=select_ppt).grid(row=1,column=2,padx=5,pady=5)

tk.Label(root, text="输出文件夹:").grid(row=2,column=0,padx=5,pady=5, sticky='e')
output_entry = tk.Entry(root, width=50); output_entry.grid(row=2,column=1)
tk.Button(root,text="选择", command=select_output).grid(row=2,column=2,padx=5,pady=5)

tk.Button(root,text="生成 PPT", command=generate, width=20).grid(row=3,column=1,pady=10)

root.mainloop()