import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import os

# 創建主應用程式
root = tk.Tk()
root.title("XPS/oxps 轉 PDF")
root.geometry("400x250")

# 保存選擇的文件和資料夾
xps_file_path = ""
output_folder_path = ""

# 選擇 XPS/oxps 文件的函數
def select_xps_file():
    global xps_file_path
    xps_file_path = filedialog.askopenfilename(filetypes=[("XPS and OXPS files", "*.xps *.oxps")])
    if xps_file_path:
        xps_label.config(text=f"選擇的文件: {os.path.basename(xps_file_path)}")
    else:
        xps_label.config(text="未選擇文件")

# 選擇保存資料夾的函數
def select_output_folder():
    global output_folder_path
    output_folder_path = filedialog.askdirectory()
    if output_folder_path:
        folder_label.config(text=f"儲存資料夾: {output_folder_path}")
    else:
        folder_label.config(text="未選擇資料夾")

# 轉換功能
def convert_xps_to_pdf():
    global dpi_entry
    if not xps_file_path or not output_folder_path:
        messagebox.showerror("錯誤", "請選擇 XPS 或 OXPS 文件和儲存資料夾")
        return

    # 獲取 DPI 值
    try:
        dpi = int(dpi_entry.get())
        if dpi <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("錯誤", "請輸入有效的 DPI 值")
        return

    # 設置 PDF 保存路徑
    pdf_filename = os.path.basename(xps_file_path).replace(".xps", ".pdf").replace(".oxps", ".pdf")
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

        messagebox.showinfo("成功", f"成功將 XPS/OXPS 轉換為 PDF：\n{pdf_file_path}")
    except Exception as e:
        messagebox.showerror("錯誤", f"轉換失敗: {str(e)}")

# 建立介面
xps_label = tk.Label(root, text="未選擇文件")
xps_label.pack(pady=10)

select_xps_button = tk.Button(root, text="選擇 XPS 或 OXPS 文件", command=select_xps_file)
select_xps_button.pack()

folder_label = tk.Label(root, text="未選擇資料夾")
folder_label.pack(pady=10)

select_folder_button = tk.Button(root, text="選擇儲存資料夾", command=select_output_folder)
select_folder_button.pack()

# 增加 DPI 輸入框
dpi_label = tk.Label(root, text="輸入 DPI (例如 300):")
dpi_label.pack(pady=5)

dpi_entry = tk.Entry(root)
dpi_entry.pack()
dpi_entry.insert(0, "150")  # 預設值 150 DPI

convert_button = tk.Button(root, text="轉換", command=convert_xps_to_pdf)
convert_button.pack(pady=20)

# 運行應用程式
root.mainloop()
