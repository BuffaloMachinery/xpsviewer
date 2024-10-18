import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import os

# 創建主應用程式
root = tk.Tk()
root.title("XPS/OXPS 轉 PDF 預覽器")
root.geometry("800x600")

# 保存選擇的文件和資料夾
xps_file_path = ""
output_folder_path = ""

# 預覽頁數
current_page = 0
total_pages = 0
zoom_level = 1.0  # 初始縮放比例

# 創建可滾動的畫布
canvas_frame = tk.Frame(root)
canvas_frame.pack(fill="both", expand=True)

canvas = tk.Canvas(canvas_frame)
scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

# 創建預覽區域
preview_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=preview_frame, anchor="nw")

# 當尺寸變更時自動調整滾動區域
def on_canvas_resize(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

preview_frame.bind("<Configure>", on_canvas_resize)

# 綁定滑鼠滾輪滾動功能
def on_mouse_wheel(event):
    if event.delta:  # Windows/Linux
        canvas.yview_scroll(-int(event.delta / 120), "units")
    else:  # macOS
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

# 綁定滑鼠滾輪事件
root.bind_all("<MouseWheel>", on_mouse_wheel)  # Windows/Linux
root.bind_all("<Button-4>", on_mouse_wheel)  # macOS 滾輪向上
root.bind_all("<Button-5>", on_mouse_wheel)  # macOS 滾輪向下

# 選擇 XPS 或 OXPS 文件的函數
def select_xps_file():
    global xps_file_path, total_pages, current_page
    xps_file_path = filedialog.askopenfilename(filetypes=[("XPS/OXPS files", "*.xps *.oxps")])
    if xps_file_path:
        xps_label.config(text=f"已選擇文件: {os.path.basename(xps_file_path)}")
        load_xps_preview()
    else:
        xps_label.config(text="未選擇 XPS/OXPS 文件")

# 加載 XPS 預覽
def load_xps_preview():
    global total_pages, current_page
    try:
        doc = fitz.open(xps_file_path)
        total_pages = doc.page_count
        current_page = 0
        display_page(current_page)
    except Exception as e:
        messagebox.showerror("錯誤", f"無法加載文件: {str(e)}")

# 顯示頁面
def display_page(page_num):
    global zoom_level
    try:
        doc = fitz.open(xps_file_path)
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom_level, zoom_level))  # 按照 zoom_level 縮放
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        img_tk = ImageTk.PhotoImage(img)
        preview_label.config(image=img_tk)
        preview_label.image = img_tk
        page_info_label.config(text=f"第 {page_num + 1} 頁 / 共 {total_pages} 頁 (縮放比例: {zoom_level:.1f}x)")

        canvas.configure(scrollregion=canvas.bbox("all"))
    except Exception as e:
        messagebox.showerror("錯誤", f"無法顯示頁面: {str(e)}")

# 下一頁
def next_page():
    global current_page
    if current_page < total_pages - 1:
        current_page += 1
        display_page(current_page)

# 上一頁
def previous_page():
    global current_page
    if current_page > 0:
        current_page -= 1
        display_page(current_page)

# 轉換為 PDF 的函數
def convert_xps_to_pdf():
    if not xps_file_path:
        messagebox.showerror("錯誤", "請選擇 XPS/OXPS 文件")
        return

    output_folder_path = filedialog.askdirectory()
    if not output_folder_path:
        messagebox.showerror("錯誤", "請選擇保存資料夾")
        return

    # 設置 PDF 保存路徑
    pdf_filename = os.path.basename(xps_file_path).replace(".xps", ".pdf").replace(".oxps", ".pdf")
    pdf_file_path = os.path.join(output_folder_path, pdf_filename)

    try:
        doc = fitz.open(xps_file_path)
        pdf_doc = fitz.open()

        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom_level, zoom_level))
            img_pdf = fitz.open()
            img_page = img_pdf.new_page(width=pix.width, height=pix.height)
            img_page.insert_image(fitz.Rect(0, 0, pix.width, pix.height), pixmap=pix)
            pdf_doc.insert_pdf(img_pdf)

        pdf_doc.save(pdf_file_path)
        pdf_doc.close()
        doc.close()

        messagebox.showinfo("成功", f"成功將 XPS/OXPS 轉換為 PDF：\n{pdf_file_path}")
    except Exception as e:
        messagebox.showerror("錯誤", f"轉換失敗: {str(e)}")

# 放大頁面
def zoom_in():
    global zoom_level
    zoom_level += 0.1
    display_page(current_page)

# 縮小頁面
def zoom_out():
    global zoom_level
    if zoom_level > 0.1:  # 確保縮放比例不低於0.1
        zoom_level -= 0.1
        display_page(current_page)

# 建立介面
xps_label = tk.Label(root, text="未選擇 XPS/OXPS 文件")
xps_label.pack(pady=10)

select_xps_button = tk.Button(root, text="選擇 XPS/OXPS 文件", command=select_xps_file)
select_xps_button.pack()

preview_label = tk.Label(preview_frame)
preview_label.pack(pady=10)

page_info_label = tk.Label(root, text="")
page_info_label.pack()

previous_button = tk.Button(root, text="上一頁", command=previous_page)
previous_button.pack(side="left", padx=20)

next_button = tk.Button(root, text="下一頁", command=next_page)
next_button.pack(side="right", padx=20)

zoom_in_button = tk.Button(root, text="放大", command=zoom_in)
zoom_in_button.pack(side="left", padx=10)

zoom_out_button = tk.Button(root, text="縮小", command=zoom_out)
zoom_out_button.pack(side="right", padx=10)

convert_button = tk.Button(root, text="轉換為 PDF", command=convert_xps_to_pdf)
convert_button.pack(pady=20)

# 運行應用程式
root.mainloop()
