
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


# ȫ���ֵ�洢OCR���
ocr_results = {}
# ����Word�ĵ�
doc = Document()

# �ص���������OCR���
def ocr_result_callback(img_path: str, results: dict):

    print(f"ʶ��ɹ���img_path: {img_path}")
   # print(results)
    if 'ocrResult' in results:
        text_result = ''.join([item['text'] for item in results['ocrResult']])
        ocr_results[img_path] = text_result
        doc.add_paragraph(text_result+'\n')
      #  print(text_result)


def pdf_to_word(pdf_path, output_path):
    global ocr_results
    ocr_results = {}

    # ��ʼ��OCR������
    ocr_manager = OcrManager(folder_path)
    ocr_manager.SetExePath(WeChatOCR_path)
    ocr_manager.SetUsrLibDir(folder_path)
    ocr_manager.SetOcrResultCallback(ocr_result_callback)
    ocr_manager.StartWeChatOCR()

    # ��PDFҳ��ת��Ϊͼ��
    images = convert_from_path(pdf_path)

    # ������ʱ�ļ��洢ͼ��
    temp_image_files = []

    for i, image in enumerate(images):
        image_path = f'temp_image_{i}.png'
        image.save(image_path)
        temp_image_files.append(image_path)
        ocr_manager.DoOCRTask(image_path)

    # �ȴ������������
    time.sleep(1)
    while ocr_manager.m_task_id.qsize() != OCR_MAX_TASK_ID:
        pass

    # ��ֹOCR����
    ocr_manager.KillWeChatOCR()

    doc.save(output_path)
    print(f"PDF converted to {output_path}")

    # ɾ����ʱͼ���ļ�
    for image_path in temp_image_files:
        os.remove(image_path)


# �������ڶ���
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
   # print("ѡ����ļ��У�", folderpath)
pdf_dir= ""
def select_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    folder_label.config(text=f"��ѡ��{folder_path}",fg='red')

def select_file():
    # ���ļ�ѡ��Ի�
    global WeChatOCR_path
    WeChatOCR_path = filedialog.askopenfilename()
    if WeChatOCR_path[-13:] == 'WeChatOCR.exe':
        WeChatOCR_label.config(text=f"��ѡ��{WeChatOCR_path}",fg='red')
    else:
        tk.messagebox.showwarning("����", "����ѡ��WeChatOCR.exe")

  #  print(filepath)

def select_pdf():
    # ���ļ�ѡ��Ի�
    global pdf_dir
    pdf_dir = filedialog.askopenfilename()

    pdf_label.config(text=f"��ѡ��{pdf_dir}",fg='red')


def confirm_selection():
    if pdf_dir[-3:] == 'pdf':
        output_path = pdf_dir[:-4] + '.docx'
        pdf_to_word(pdf_dir, output_path)
        font_style = ("Arial", 20, "bold")
        out_label.config(text=f"�����{output_path}",font=font_style)
    else:
        tk.messagebox.showwarning("����", "����ѡ��һ��PDF�ļ�")




folder_label = tk.Label(window, text="δѡ��΢�������ļ���")
folder_label.pack(pady=10)

select_button = tk.Button(window, text="1��ѡ��wechat�ļ���", command=select_folder)
select_button.pack(pady=5)


Label= tk.Label(window, text="����C:\Program Files (x86)\Tencent\WeChat\[3.9.0.28]")
Label.pack()

WeChatOCR_label = tk.Label(window, text="δѡ��WeChatOCR")
WeChatOCR_label.pack(pady=10)

select_OCRbutton = tk.Button(window, text="2��ѡ��WeChatOCR.exe", command=select_file)
select_OCRbutton.pack(pady=10)

Label= tk.Label(window, text=str('����C:/Users/unicom/AppData/Roaming/Tencent/WXWork/upgrade/4.1.27.6032/WeChatOCR/WeChatOCR.exe'))
Label.pack(pady=10)


pdf_label = tk.Label(window, text="δѡ��pdf�ļ�")
pdf_label.pack(pady=10)

pdf_button = tk.Button(window, text="3��ѡ��pdf�ļ�", command=select_pdf)
pdf_button.pack(pady=10)

confirm_button = tk.Button(window, text="ȷ��", command=confirm_selection)
confirm_button.pack(pady=5)

out_label = tk.Label(window, text="�������")
out_label.pack(pady=10)

# ������ѭ��
window.mainloop()
