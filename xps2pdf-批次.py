import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import os

# 創建主應用程式
root = tk.Tk()
root.title("XPS/oxps 批次轉 PDF")
root.geometry("400x250")

# 保存選擇的資料夾
input_folder_path = ""
output_folder_path = ""

# 選擇包含 XPS/oxps 文件的資料夾
def select_input_folder():
    global input_folder_path
    input_folder_path = filedialog.askdirectory()
    if input_folder_path:
        input_folder_label.config(text=f"來源資料夾: {input_folder_path}")
    else:
        input_folder_label.config(text="未選擇來源資料夾")

# 選擇保存 PDF 文件的資料夾
def select_output_folder():
    global output_folder_path
    output_folder_path = filedialog.askdirectory()
    if output_folder_path:
        output_folder_label.config(text=f"儲存資料夾: {output_folder_path}")
    else:
        output_folder_label.config(text="未選擇儲存資料夾")

# 轉換功能
def convert_batch_xps_to_pdf():
    global dpi_entry
    if not input_folder_path or not output_folder_path:
        messagebox.showerror("錯誤", "請選擇來源資料夾和儲存資料夾")
        return

    # 獲取 DPI 值
    try:
        dpi = int(dpi_entry.get())
        if dpi <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("錯誤", "請輸入有效的 DPI 值")
        return

    # 找到來源資料夾中的所有 .xps 和 .oxps 文件
    files = [f for f in os.listdir(input_folder_path) if f.lower().endswith(('.xps', '.oxps'))]
    
    if not files:
        messagebox.showerror("錯誤", "來源資料夾中沒有 .xps 或 .oxps 文件")
        return

    failed_files = []

    for file_name in files:
        xps_file_path = os.path.join(input_folder_path, file_name)
        pdf_filename = file_name.replace(".xps", ".pdf").replace(".oxps", ".pdf")
        pdf_file_path = os.path.join(output_folder_path, pdf_filename)

        try:
            # 開始轉換
            doc = fitz.open(xps_file_path)
            pdf_doc = fitz.open()

            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                # 使用用戶指定的 DPI
                pix = page.get_pixmap(dpi=dpi)
                img_pdf = fitz.open()
                img_page = img_pdf.new_page(width=pix.width, height=pix.height)
                img_page.insert_image(fitz.Rect(0, 0, pix.width, pix.height), pixmap=pix)
                pdf_doc.insert_pdf(img_pdf)

            # 保存 PDF 文件
            pdf_doc.save(pdf_file_path)
            pdf_doc.close()
            doc.close()

        except Exception as e:
            failed_files.append(file_name)
            continue

    if failed_files:
        messagebox.showerror("錯誤", f"以下文件轉換失敗: {', '.join(failed_files)}")
    else:
        messagebox.showinfo("成功", f"所有文件成功轉換為 PDF。")

# 建立介面
input_folder_label = tk.Label(root, text="未選擇來源資料夾")
input_folder_label.pack(pady=10)

select_input_folder_button = tk.Button(root, text="選擇來源資料夾", command=select_input_folder)
select_input_folder_button.pack()

output_folder_label = tk.Label(root, text="未選擇儲存資料夾")
output_folder_label.pack(pady=10)

select_output_folder_button = tk.Button(root, text="選擇儲存資料夾", command=select_output_folder)
select_output_folder_button.pack()

# 增加 DPI 輸入框
dpi_label = tk.Label(root, text="輸入 DPI (例如 150):")
dpi_label.pack(pady=5)

dpi_entry = tk.Entry(root)
dpi_entry.pack()
dpi_entry.insert(0, "150")  # 預設值 150 DPI

convert_button = tk.Button(root, text="批次轉換", command=convert_batch_xps_to_pdf)
convert_button.pack(pady=20)

# 運行應用程式
root.mainloop()
