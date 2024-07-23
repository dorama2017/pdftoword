
# -*- coding: gb2312 -*-
import os
import time
from pdf2image import convert_from_path
from docx import Document
import tkinter as tk
from tkinter import filedialog

from wechat_ocr.ocr_manager import OcrManager, OCR_MAX_TASK_ID

#wechat_ocr_dir = r"C:\Users\unicom\AppData\Roaming\Tencent\WXWork\upgrade\4.1.27.6032\WeChatOCR\WeChatOCR.exe"
#wechat_dir = r"C:\Program Files (x86)\Tencent\WeChat\[3.9.0.28]"


# 全局字典存储OCR结果
ocr_results = {}
# 创建Word文档
doc = Document()

# 回调函数处理OCR结果
def ocr_result_callback(img_path: str, results: dict):

    print(f"识别成功，img_path: {img_path}")
   # print(results)
    if 'ocrResult' in results:
        text_result = ''.join([item['text'] for item in results['ocrResult']])
        ocr_results[img_path] = text_result
        doc.add_paragraph(text_result+'\n')
      #  print(text_result)


def pdf_to_word(pdf_path, output_path):
    global ocr_results
    ocr_results = {}

    # 初始化OCR管理器
    ocr_manager = OcrManager(folder_path)
    ocr_manager.SetExePath(WeChatOCR_path)
    ocr_manager.SetUsrLibDir(folder_path)
    ocr_manager.SetOcrResultCallback(ocr_result_callback)
    ocr_manager.StartWeChatOCR()

    # 将PDF页面转换为图像
    images = convert_from_path(pdf_path)

    # 创建临时文件存储图像
    temp_image_files = []

    for i, image in enumerate(images):
        image_path = f'temp_image_{i}.png'
        image.save(image_path)
        temp_image_files.append(image_path)
        ocr_manager.DoOCRTask(image_path)

    # 等待所有任务完成
    time.sleep(1)
    while ocr_manager.m_task_id.qsize() != OCR_MAX_TASK_ID:
        pass

    # 终止OCR服务
    ocr_manager.KillWeChatOCR()

    doc.save(output_path)
    print(f"PDF converted to {output_path}")

    # 删除临时图像文件
    for image_path in temp_image_files:
        os.remove(image_path)


# 创建窗口对象
window = tk.Tk()
window.title("PDF To Word Converter")
#screen_width = window.winfo_screenwidth()
#screen_height = window.winfo_screenheight()
#x = int(screen_width / 2 - 1000 / 2)
#y = int(screen_height / 2 - 400 / 2)
size = '{}x{}+{}+{}'.format(1000, 500, 0, 0)
window.geometry(size)

folder_path = "C:\Program Files (x86)\Tencent\WeChat\[3.9.0.28]"
WeChatOCR_path =r"C:\Users\unicom\AppData\Roaming\Tencent\WXWork\WeChatOCR\1.0.1.20\WeChatOCR\WeChatOCR.exe"
   # print("选择的文件夹：", folderpath)
pdf_dir= ""
def select_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    folder_label.config(text=f"已选中{folder_path}",fg='red')

def select_file():
    # 打开文件选择对话
    global WeChatOCR_path
    WeChatOCR_path = filedialog.askopenfilename()
    if WeChatOCR_path[-13:] == 'WeChatOCR.exe':
        WeChatOCR_label.config(text=f"已选中{WeChatOCR_path}",fg='red')
    else:
        tk.messagebox.showwarning("警告", "请先选择WeChatOCR.exe")

  #  print(filepath)

def select_pdf():
    # 打开文件选择对话
    global pdf_dir
    pdf_dir = filedialog.askopenfilename()

    pdf_label.config(text=f"已选中{pdf_dir}",fg='red')


def confirm_selection():
    if pdf_dir[-3:] == 'pdf':
        output_path = pdf_dir[:-4] + '.docx'
        pdf_to_word(pdf_dir, output_path)
        font_style = ("Arial", 20, "bold")
        out_label.config(text=f"已输出{output_path}",font=font_style)
    else:
        tk.messagebox.showwarning("警告", "请先选择一个PDF文件")




folder_label = tk.Label(window, text="未选中微信所在文件夹")
folder_label.pack(pady=10)

select_button = tk.Button(window, text="1、选择wechat文件夹", command=select_folder)
select_button.pack(pady=5)


Label= tk.Label(window, text="例：C:\Program Files (x86)\Tencent\WeChat\[3.9.0.28]")
Label.pack()

WeChatOCR_label = tk.Label(window, text="未选中WeChatOCR")
WeChatOCR_label.pack(pady=10)

select_OCRbutton = tk.Button(window, text="2、选择WeChatOCR.exe", command=select_file)
select_OCRbutton.pack(pady=10)

Label= tk.Label(window, text=str('例：C:/Users/unicom/AppData/Roaming/Tencent/WXWork/upgrade/4.1.27.6032/WeChatOCR/WeChatOCR.exe'))
Label.pack(pady=10)


pdf_label = tk.Label(window, text="未选中pdf文件")
pdf_label.pack(pady=10)

pdf_button = tk.Button(window, text="3、选择pdf文件", command=select_pdf)
pdf_button.pack(pady=10)

confirm_button = tk.Button(window, text="确定", command=confirm_selection)
confirm_button.pack(pady=5)

out_label = tk.Label(window, text="暂无输出")
out_label.pack(pady=10)

# 启动主循环
window.mainloop()
