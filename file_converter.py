import os
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import img2pdf
from PIL import Image
import threading
import fitz
import openpyxl
import xlrd
from openpyxl import Workbook
import traceback
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

logging.getLogger("pikepdf").setLevel(logging.ERROR)
Image.MAX_IMAGE_PIXELS = None
os.environ["PYDEVD_DISABLE_FILE_VALIDATION"] = "1"

def get_real_desktop_path():
    try:
        from win32com.shell import shell, shellcon
        return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
    except:
        return os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")


def check_pdf_type(pdf_path):
    try:
        # 使用PyMuPDF提取文本
        doc = fitz.open(pdf_path)
        total_text = ""
        for page_num in range(min(3, len(doc))):
            page = doc[page_num]
            text = page.get_text().strip()
            total_text += text
            if len(total_text) > 100:
                doc.close()
                return "文本型"
        doc.close()
        return "扫描型"
    except Exception as e:
        print(f"PDF类型检测异常: {str(e)}")
        return "扫描型"

def clean_temp_files(file_list):
    for f in file_list:
        if os.path.exists(f):
            try:
                os.remove(f)
            except:
                pass

def text_pdf_to_word(input_file, output_file, progress_callback=None):
    """
    文本型PDF转Word - 使用PyMuPDF提取文本，然后用python-docx创建Word
    """
    try:
        # 打开PDF文件
        pdf_doc = fitz.open(input_file)
        total_pages = len(pdf_doc)

        # 创建Word文档
        doc = Document()

        # 设置页面为A4大小
        section = doc.sections[0]
        section.page_height = Cm(29.7)  # A4高度
        section.page_width = Cm(21)  # A4宽度

        if progress_callback:
            progress_callback(20)

        # 提取每一页文本
        for page_num in range(total_pages):
            page = pdf_doc[page_num]
            text = page.get_text()

            if text.strip():
                # 添加段落
                paragraph = doc.add_paragraph(text)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                # 如果不是最后一页，添加分页符
                if page_num < total_pages - 1:
                    doc.add_page_break()

            # 更新进度
            if progress_callback:
                progress = 20 + (page_num + 1) / total_pages * 70
                progress_callback(progress)

        pdf_doc.close()

        # 保存Word文档
        doc.save(output_file)

        if progress_callback:
            progress_callback(100)
        return True, f"文本型PDF转换完成！共提取{total_pages}页文本"

    except Exception as e:
        error_msg = f"文本型PDF转换失败：{str(e)}"
        return False, error_msg


def scan_pdf_to_word(input_file, output_file, progress_callback=None):
    """
    扫描型PDF转Word - 每页转为图片插入Word（使用python-docx）
    """
    temp_files = []
    try:
        # 1. 使用PyMuPDF将PDF每页转为图片
        pdf_doc = fitz.open(input_file)
        total_pages = len(pdf_doc)
        img_paths = []

        # 创建临时目录
        temp_dir = os.path.join(os.environ.get("TEMP", ""), "pdf_converter_temp")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        # 转换每一页为图片
        for idx in range(total_pages):
            page = pdf_doc[idx]
            # 计算DPI使图片清晰
            zoom = 2.0
            matrix = fitz.Matrix(zoom, zoom)

            # 渲染页面为图片
            pix = page.get_pixmap(matrix=matrix)
            img_path = os.path.join(temp_dir, f"temp_page_{idx + 1}.png")
            pix.save(img_path)
            img_paths.append(img_path)
            temp_files.append(img_path)

            # 更新进度
            if progress_callback:
                progress = (idx + 1) / total_pages * 50
                progress_callback(progress)

        pdf_doc.close()

        if progress_callback:
            progress_callback(55)

        # 2. 创建Word文档并插入图片
        doc = Document()

        # 设置页面为A4大小
        section = doc.sections[0]
        section.page_height = Cm(29.7)  # A4高度
        section.page_width = Cm(21)  # A4宽度

        # 设置页边距
        section.top_margin = Cm(1.27)  # 上边距
        section.bottom_margin = Cm(1.27)  # 下边距
        section.left_margin = Cm(1.27)  # 左边距
        section.right_margin = Cm(1.27)  # 右边距

        # 计算图片大小（A4页面宽度减去左右边距）
        image_width = section.page_width - section.left_margin - section.right_margin

        # 插入图片到Word
        for i, img_path in enumerate(img_paths):
            # 添加图片
            paragraph = doc.add_paragraph()
            run = paragraph.add_run()

            # 插入图片，设置宽度为页面宽度
            run.add_picture(img_path, width=image_width)

            # 居中图片
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 如果不是最后一页，添加分页符
            if i < len(img_paths) - 1:
                doc.add_page_break()

        if progress_callback:
            progress_callback(85)

        # 保存Word文档
        doc.save(output_file)

        # 清理临时文件
        clean_temp_files(temp_files)

        if progress_callback:
            progress_callback(100)
        return True, f"扫描型PDF转换完成！共{total_pages}页已转为图片插入Word"

    except Exception as e:
        # 清理临时文件
        clean_temp_files(temp_files)
        error_msg = f"扫描型PDF转换失败：{str(e)}"
        return False, error_msg

def img_to_pdf(input_files, output_file, progress_callback=None):
    try:
        img_list = []
        total_imgs = len(input_files)

        for idx, img_path in enumerate(input_files):
            with Image.open(img_path) as img:
                # 转换RGBA为RGB
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                img_list.append(img)

            if progress_callback:
                progress = (idx + 1) / total_imgs * 90
                progress_callback(progress)

        # 保存为PDF
        if img_list:
            img_list[0].save(
                output_file,
                "PDF",
                save_all=True,
                append_images=img_list[1:] if len(img_list) > 1 else [],
                quality=95
            )

        if progress_callback:
            progress_callback(100)
        return True, f"图片转PDF完成！共转换{len(img_list)}张图片"
    except Exception as e:
        return False, f"图片转PDF失败：{str(e)}"


def word_to_pdf(input_file, output_file, progress_callback=None):
    """
    Word转PDF - 使用PyMuPDF将Word先转图片再合成PDF
    """
    try:
        # 先将Word转为图片，再将图片转为PDF
        temp_dir = os.path.join(os.environ.get("TEMP", ""), "word_to_pdf_temp")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        try:
            from docx import Document
            doc = Document(input_file)

            # 创建PDF文档
            pdf_doc = fitz.open()

            # 对于每个段落，创建一个PDF页面
            for para in doc.paragraphs:
                if para.text.strip():
                    # 创建新页面
                    page = pdf_doc.new_page(width=595, height=842)

                    # 插入文本
                    text = para.text
                    page.insert_text((50, 50), text, fontsize=12)

            if len(pdf_doc) == 0:
                pdf_doc.new_page(width=595, height=842)

            pdf_doc.save(output_file)
            pdf_doc.close()

            if progress_callback:
                progress_callback(100)
            return True, "Word转PDF完成（文本内容转换）"

        except Exception as e:
            # 如果上述方法失败，使用图片转换方法
            return word_to_img_then_pdf(input_file, output_file, progress_callback)

    except Exception as e:
        error_msg = f"Word转PDF失败：{str(e)}"
        return False, error_msg


def word_to_img_then_pdf(input_file, output_file, progress_callback=None):
    """
    备用方法：先将Word转为图片，再将图片转为PDF
    """
    try:
        temp_dir = os.path.join(os.environ.get("TEMP", ""), "word_temp")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        # 使用word_to_img函数先将Word转为图片
        success, msg = word_to_img(input_file, temp_dir, progress_callback)

        if not success:
            return False, f"Word转PDF失败：{msg}"

        # 获取生成的图片文件
        img_files = []
        for f in os.listdir(temp_dir):
            if f.endswith(('.png', '.jpg', '.jpeg')):
                img_files.append(os.path.join(temp_dir, f))

        if not img_files:
            return False, "Word转PDF失败：未生成图片文件"

        # 将图片转为PDF
        img_list = []
        for img_path in img_files:
            with Image.open(img_path) as img:
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                img_list.append(img)

        if img_list:
            img_list[0].save(
                output_file,
                "PDF",
                save_all=True,
                append_images=img_list[1:] if len(img_list) > 1 else [],
                quality=95
            )

        # 清理临时文件
        for f in img_files:
            if os.path.exists(f):
                os.remove(f)
        if os.path.exists(temp_dir):
            os.rmdir(temp_dir)

        return True, f"Word转PDF完成！共转换{len(img_files)}页"
    except Exception as e:
        return False, f"Word转PDF失败：{str(e)}"


def word_to_img(input_file, output_dir, progress_callback=None):
    """
    Word转图片 - 使用python-docx读取内容，然后渲染为图片
    """
    try:
        from docx import Document
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        doc = Document(input_file)
        paragraphs = list(doc.paragraphs)
        total_paras = len(paragraphs)

        if total_paras == 0:
            img = Image.new('RGB', (800, 600), color='white')
            img_path = os.path.join(output_dir, "word_page_1.png")
            img.save(img_path)

            if progress_callback:
                progress_callback(100)
            return True, f"Word转图片完成！保存至：{output_dir}"
        for idx, para in enumerate(paragraphs):
            if para.text.strip():
                img = Image.new('RGB', (800, 200), color='white')

                # 这里简化处理，实际应该使用更复杂的文本渲染
                # 由于PIL的文本渲染功能有限，这里只做简单示例
                from PIL import ImageDraw, ImageFont

                draw = ImageDraw.Draw(img)
                try:
                    font = ImageFont.truetype("arial.ttf", 14)
                except:
                    font = ImageFont.load_default()
                draw.text((10, 10), para.text, fill='black', font=font)

                img_path = os.path.join(output_dir, f"word_page_{idx + 1}.png")
                img.save(img_path)
            if progress_callback:
                progress = (idx + 1) / total_paras * 100
                progress_callback(progress)

        return True, f"Word转图片完成！共生成{total_paras}张图片，保存至：{output_dir}"
    except Exception as e:
        error_msg = f"Word转图片失败：{str(e)}"
        return False, error_msg


def pdf_to_img(input_file, output_dir, progress_callback=None):
    try:
        pdf_doc = fitz.open(input_file)
        total_pages = len(pdf_doc)

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        for page_num in range(total_pages):
            page = pdf_doc[page_num]
            pix = page.get_pixmap(dpi=150)
            img_path = os.path.join(output_dir, f"pdf_page_{page_num + 1}.png")
            pix.save(img_path)

            if progress_callback:
                progress = (page_num + 1) / total_pages * 100
                progress_callback(progress)

        pdf_doc.close()
        return True, f"PDF转图片完成！共生成{total_pages}张图片，保存至：{output_dir}"
    except Exception as e:
        return False, f"PDF转图片失败：{str(e)}"


def xls_to_word(input_file, output_file, progress_callback=None):
    try:
        data = []
        sheet_names = []

        if input_file.endswith(".xls"):
            wb = xlrd.open_workbook(input_file)
            sheet_names = wb.sheet_names()
            for sheet_name in sheet_names:
                ws = wb.sheet_by_name(sheet_name)
                sheet_data = []
                for row_idx in range(ws.nrows):
                    row_data = [ws.cell_value(row_idx, col_idx) for col_idx in range(ws.ncols)]
                    sheet_data.append(row_data)
                data.append((sheet_name, sheet_data))
            wb.release_resources()
        else:
            # 处理.xlsx文件
            wb = openpyxl.load_workbook(input_file, data_only=True)
            sheet_names = wb.sheetnames
            for sheet_name in sheet_names:
                ws = wb[sheet_name]
                sheet_data = []
                for row in ws.iter_rows(values_only=True):
                    sheet_data.append([cell if cell is not None else "" for cell in row])
                data.append((sheet_name, sheet_data))
            wb.close()

        if progress_callback:
            progress_callback(40)

        # 使用python-docx创建Word文档
        doc = Document()

        # 设置页面为A4大小
        section = doc.sections[0]
        section.page_height = Cm(29.7)
        section.page_width = Cm(21)

        # 添加标题
        doc.add_heading('Excel数据转换结果', 0)

        # 逐工作表处理
        total_sheets = len(data)
        for sheet_idx, (sheet_name, sheet_data) in enumerate(data):
            # 添加工作表标题
            doc.add_heading(f'工作表：{sheet_name}', level=1)

            if sheet_data:
                # 创建表格
                table = doc.add_table(rows=len(sheet_data), cols=len(sheet_data[0]))
                table.style = 'Table Grid'

                # 填充表格内容
                for row_idx, row_data in enumerate(sheet_data):
                    row_cells = table.rows[row_idx].cells
                    for col_idx, cell_value in enumerate(row_data):
                        row_cells[col_idx].text = str(cell_value)

            # 添加分页符（如果不是最后一个工作表）
            if sheet_idx < total_sheets - 1:
                doc.add_page_break()

            # 更新进度
            if progress_callback:
                progress = 40 + (sheet_idx + 1) / total_sheets * 50
                progress_callback(progress)

        # 保存Word文档
        doc.save(output_file)

        if progress_callback:
            progress_callback(100)
        return True, f"XLS转Word完成！共处理{total_sheets}个工作表，{len(data[0][1]) if data else 0}行数据"
    except Exception as e:
        error_msg = f"XLS转Word失败：{str(e)}"
        return False, error_msg
    #UI
def select_file(converter_type, file_path, root):
    file_path.set("")
    title_map = {
        "pdf2word": "选择要转换的PDF文件",
        "word2pdf": "选择要转换的Word文件",
        "img2pdf": "选择要转换的图片文件",
        "word2img": "选择要转换的Word文件",
        "pdf2img": "选择要转换的PDF文件",
        "xls2word": "选择要转换的Excel文件"
    }
    type_map = {
        "pdf2word": [("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        "word2pdf": [("Word文件", "*.docx;*.doc"), ("所有文件", "*.*")],
        "img2pdf": [("图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"), ("所有文件", "*.*")],
        "word2img": [("Word文件", "*.docx;*.doc"), ("所有文件", "*.*")],
        "pdf2img": [("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        "xls2word": [("Excel文件", "*.xls;*.xlsx"), ("所有文件", "*.*")]
    }

    files = filedialog.askopenfilenames(title=title_map[converter_type], filetypes=type_map[converter_type])
    if files:
        root.file_list = list(files)
        display = f"✅ 已选中：\n{files[0]}" if len(files) == 1 else f"✅ 已选中 {len(files)} 个文件"
        file_path.set(display)


def update_progress(value, progress_var, root):
    progress_var.set(value)
    root.update_idletasks()


def convert_thread(converter_type, input_files, new_name, root, progress_var, convert_btn, name_entry, file_path):
    desktop = get_real_desktop_path()
    if not os.path.exists(desktop) or not os.access(desktop, os.W_OK):
        desktop = os.path.dirname(sys.executable) if hasattr(sys, 'frozen') else os.getcwd()
        root.after(0, lambda: messagebox.showwarning("权限提示", f"桌面不可写，保存到：\n{desktop}"))

    success = False
    msg = ""
    output = ""

    try:
        if converter_type == "pdf2word":
            selected_pdf = input_files[0]

            # 检测PDF类型
            pdf_type = check_pdf_type(selected_pdf)
            root.after(0, lambda: messagebox.showinfo("PDF类型检测", f"检测到PDF类型：{pdf_type}\n开始转换..."))

            output = os.path.join(desktop, f"{new_name}.docx")

            if pdf_type == "文本型":
                success, msg = text_pdf_to_word(
                    selected_pdf,
                    output,
                    lambda v: update_progress(v, progress_var, root)
                )
            else:
                success, msg = scan_pdf_to_word(
                    selected_pdf,
                    output,
                    lambda v: update_progress(v, progress_var, root)
                )

        elif converter_type == "word2pdf":
            selected_word = input_files[0]
            output = os.path.join(desktop, f"{new_name}.pdf")
            success, msg = word_to_pdf(
                selected_word,
                output,
                lambda v: update_progress(v, progress_var, root)
            )

        elif converter_type == "img2pdf":
            output = os.path.join(desktop, f"{new_name}.pdf")
            success, msg = img_to_pdf(
                input_files,
                output,
                lambda v: update_progress(v, progress_var, root)
            )

        elif converter_type == "word2img":
            selected_word = input_files[0]
            output = os.path.join(desktop, new_name)
            success, msg = word_to_img(
                selected_word,
                output,
                lambda v: update_progress(v, progress_var, root)
            )

        elif converter_type == "pdf2img":
            selected_pdf = input_files[0]
            output = os.path.join(desktop, new_name)
            success, msg = pdf_to_img(
                selected_pdf,
                output,
                lambda v: update_progress(v, progress_var, root)
            )

        elif converter_type == "xls2word":
            selected_xls = input_files[0]
            output = os.path.join(desktop, f"{new_name}.docx")
            success, msg = xls_to_word(
                selected_xls,
                output,
                lambda v: update_progress(v, progress_var, root)
            )

        else:
            msg = "无效转换类型"

    except Exception as e:
        msg = f"执行异常：{type(e).__name__} - {str(e)}\n\n详细错误：{traceback.format_exc()}"

    root.after(0, lambda: finish_convert(success, msg, output, root, convert_btn, progress_var, name_entry, file_path))


def finish_convert(success, msg, output, root, convert_btn, progress_var, name_entry, file_path):
    convert_btn.config(state=tk.NORMAL, text="🚀 开始转换")
    progress_var.set(0)

    if success:
        messagebox.showinfo("转换成功 🎉", f"{msg}\n\n文件保存位置：\n{output}")
        name_entry.delete(0, tk.END)
        file_path.set("")
        root.file_list = []
    else:
        error_msg = f"转换失败 ❌\n\n{msg}\n\n目标路径：{output}"
        messagebox.showerror("错误", error_msg)


def start_convert(var, root, progress_var, convert_btn, name_entry, file_path):
    converter_type = var.get()
    input_files = getattr(root, "file_list", [])
    new_name = name_entry.get().strip()

    if not input_files:
        messagebox.showwarning("提示", "请先选择文件！")
        return
    if not new_name:
        messagebox.showwarning("提示", "请输入文件名！")
        return
    if any(char in new_name for char in r'\/:*?"<>|'):
        messagebox.showwarning("提示", "文件名不能包含：\\ / : * ? \" < > |")
        return

    # 检查文件是否存在
    for file in input_files:
        if not os.path.exists(file):
            messagebox.showerror("错误", f"选中的文件不存在：\n{file}")
            return

    convert_btn.config(state=tk.DISABLED, text="⏳ 转换中...")
    progress_var.set(0)

    t = threading.Thread(
        target=convert_thread,
        args=(converter_type, input_files, new_name, root, progress_var, convert_btn, name_entry, file_path),
        daemon=True
    )
    t.start()

    #主界面
if __name__ == "__main__":
    root = tk.Tk()
    root.title("📄 FMgaic")
    root.geometry("900x580")
    root.resizable(False, False)
    root.configure(bg="#f8f9fa")


    # 窗口居中
    def center_win():
        root.update_idletasks()
        w = root.winfo_width()
        h = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (w // 2)
        y = (root.winfo_screenheight() // 2) - (h // 2)
        root.geometry(f"{w}x{h}+{x}+{y}")


    center_win()

    # 变量初始化
    progress_var = tk.DoubleVar(value=0)
    var = tk.StringVar(value="pdf2word")
    file_path = tk.StringVar(value="")
    root.file_list = []

    # 样式配置
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Modern.TButton", font=("微软雅黑", 11, "bold"),
                    background="#0d6efd", foreground="white", relief=tk.FLAT, padding=8)
    style.map("Modern.TButton",
              background=[("active", "#0b5ed7"), ("disabled", "#cccccc")],
              foreground=[("active", "white"), ("disabled", "#666666")])
    style.configure("Big.TButton", font=("微软雅黑", 14, "bold"),
                    background="#0d6efd", foreground="white", relief=tk.FLAT, padding=(30, 12))
    style.configure("Modern.TEntry", font=("微软雅黑", 11), padding=8,
                    relief=tk.FLAT, fieldbackground="white", foreground="#333333")
    style.configure("Modern.TRadiobutton", font=("微软雅黑", 11),
                    foreground="#0d6efd", background="#f8f9fa", padding=6)

    # 标题 - 居中
    title_frame = tk.Frame(root, bg="#f8f9fa")
    title_frame.pack(fill=tk.X, pady=(10, 5), padx=40)
    title_label = tk.Label(title_frame, text="FMgaic",
                           font=("微软雅黑", 22, "bold"), fg="#0d6efd", bg="#f8f9fa")
    title_label.pack(expand=True)

    # 转换类型
    type_card = tk.Frame(root, bg="white", bd=0, relief=tk.FLAT)
    type_card.pack(fill=tk.X, padx=40, pady=8)
    type_title = tk.Label(type_card, text="📌 选择转换类型",
                          font=("微软雅黑", 12, "bold"), fg="#0d6efd", bg="white",
                          anchor="w", padx=20, pady=8)
    type_title.pack(fill=tk.X)

    # 创建两个Frame来分别放置两行选项
    r_frame = tk.Frame(type_card, bg="white", padx=20, pady=5)
    r_frame.pack(fill=tk.X)

    row1 = tk.Frame(r_frame, bg="white")
    row1.pack(fill=tk.X, pady=2)

    for i, (text, value) in enumerate(
            [("PDF → Word", "pdf2word"), ("Word → PDF", "word2pdf"), ("图片 → PDF", "img2pdf")]):
        cell_frame = tk.Frame(row1, bg="white")
        cell_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)

        radio = ttk.Radiobutton(cell_frame, text=text, variable=var, value=value,
                                style="Modern.TRadiobutton")
        radio.pack(fill=tk.X, padx=5, pady=3)

    row2 = tk.Frame(r_frame, bg="white")
    row2.pack(fill=tk.X, pady=2)

    for i, (text, value) in enumerate(
            [("Word → 图片", "word2img"), ("PDF → 图片", "pdf2img"), ("XLS → Word", "xls2word")]):
        cell_frame = tk.Frame(row2, bg="white")
        cell_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)

        radio = ttk.Radiobutton(cell_frame, text=text, variable=var, value=value,
                                style="Modern.TRadiobutton")
        radio.pack(fill=tk.X, padx=5, pady=3)

    # 文件选择
    file_card = tk.Frame(root, bg="white", bd=0, relief=tk.FLAT)
    file_card.pack(fill=tk.X, padx=40, pady=8)
    file_title = tk.Label(file_card, text="📂 选择文件",
                          font=("微软雅黑", 12, "bold"), fg="#0d6efd", bg="white",
                          anchor="w", padx=20, pady=8)
    file_title.pack(fill=tk.X)
    file_inner = tk.Frame(file_card, bg="white", padx=20, pady=5)
    file_inner.pack(fill=tk.X)
    file_entry = ttk.Entry(file_inner, textvariable=file_path, state="readonly",
                           style="Modern.TEntry", width=80)
    file_entry.pack(side=tk.LEFT, padx=(0, 15), fill=tk.X, expand=True)
    select_btn = ttk.Button(file_inner, text="Go",
                            command=lambda: select_file(var.get(), file_path, root),
                            style="Modern.TButton", width=10)
    select_btn.pack(side=tk.RIGHT)

    # 文件名
    name_card = tk.Frame(root, bg="white", bd=0, relief=tk.FLAT)
    name_card.pack(fill=tk.X, padx=40, pady=8)
    name_title = tk.Label(name_card, text="✏️ 自定义文件名/目录名",
                          font=("微软雅黑", 12, "bold"), fg="#0d6efd", bg="white",
                          anchor="w", padx=20, pady=8)
    name_title.pack(fill=tk.X)
    name_inner = tk.Frame(name_card, bg="white", padx=20, pady=5)
    name_inner.pack(fill=tk.X)
    name_entry = ttk.Entry(name_inner, style="Modern.TEntry", width=80)
    name_entry.pack(side=tk.LEFT, padx=(0, 15), fill=tk.X, expand=True)

    # 进度条
    progress_frame = tk.Frame(root, bg="#f8f9fa")
    progress_frame.pack(fill=tk.X, padx=40, pady=15)
    progress_bar = ttk.Progressbar(progress_frame, variable=progress_var,
                                   maximum=100, length=820)
    progress_bar.pack(fill=tk.X)

    # 转换按钮
    convert_frame = tk.Frame(root, bg="#f8f9fa")
    convert_frame.pack(pady=10)
    convert_btn = ttk.Button(convert_frame, text="🚀 开始转换",
                             style="Big.TButton", width=30)
    convert_btn.pack()
    convert_btn.config(command=lambda: start_convert(var, root, progress_var,
                                                     convert_btn, name_entry, file_path))

    root.mainloop()