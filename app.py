import os
import win32print
import win32api
import win32ui
import win32com.client
import winsound
from flask import Flask, request, redirect, url_for, render_template, send_from_directory, send_file, abort, after_this_request
from PIL import Image, ImageWin
import pythoncom
import time
import subprocess
from datetime import datetime
import io

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
SCAN_FOLDER = os.path.join(BASE_DIR, 'scans')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'xlsx', 'doc', 'xls', 'jpg', 'jpeg', 'png', 'bmp', 'txt'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(SCAN_FOLDER, exist_ok=True)
os.system("taskkill /f /im winword.exe >nul 2>&1")

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def print_image_to_printer(image_path, printer_name):
    image = Image.open(image_path)
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)

    printable_width = hdc.GetDeviceCaps(110)
    printable_height = hdc.GetDeviceCaps(111)
    total_width = hdc.GetDeviceCaps(8)
    total_height = hdc.GetDeviceCaps(10)

    img_w, img_h = image.size
    scale = min(printable_width / img_w, printable_height / img_h)
    scaled_w = int(img_w * scale)
    scaled_h = int(img_h * scale)

    x1 = int((total_width - scaled_w) / 2)
    y1 = int((total_height - scaled_h) / 2)
    x2 = x1 + scaled_w
    y2 = y1 + scaled_h

    hdc.StartDoc(image_path)
    hdc.StartPage()
    dib = ImageWin.Dib(image)
    dib.draw(hdc.GetHandleOutput(), (x1, y1, x2, y2))
    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()

def print_pdf_to_printer(pdf_path, printer_name):
    sumatra_path = r"C:\Users\chenjunqi\AppData\Local\SumatraPDF\SumatraPDF.exe"  # 修改为你的安装路径
    if not os.path.exists(sumatra_path):
        raise RuntimeError("未找到 SumatraPDF.exe，请检查路径")
    cmd = [sumatra_path, "-print-to", printer_name, pdf_path]
    subprocess.run(cmd, shell=False)
    #cmd = f'"{sumatra_path}" -print-to "{printer_name}" "{pdf_path}"'
    #os.system(cmd)

def print_word_to_printer(doc_path, printer_name):
    pythoncom.CoInitialize()
    time.sleep(0.5)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.ActivePrinter = printer_name
    doc = word.Documents.Open(doc_path)
    doc.PrintOut()
    doc.Close(False)
    word.Quit()

def print_excel_to_printer(xls_path, printer_name):
    import pythoncom
    import time
    pythoncom.CoInitialize()
    time.sleep(0.5)

    pdf_path = xls_path.rsplit('.', 1)[0] + '.pdf'
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(xls_path)
        wb.ExportAsFixedFormat(0, pdf_path)  # 0 = PDF
        wb.Close(False)
    finally:
        excel.Quit()

    print_pdf_to_printer(pdf_path, printer_name)

    if os.path.exists(pdf_path):
        os.remove(pdf_path)



@app.route('/')
def index():
    printers = [p[2] for p in win32print.EnumPrinters(2)]
    uploaded_files = os.listdir(UPLOAD_FOLDER)
    return render_template('index.html', printers=printers, files=uploaded_files)

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('file')
    if not file or not allowed_file(file.filename):
        return '文件类型不支持', 400

    filename = file.filename
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    ext = filename.rsplit('.', 1)[-1].lower()
    printer_name = request.form.get('printer')

    try:
        if ext in {'png', 'jpg', 'jpeg', 'bmp'}:
            print_image_to_printer(filepath, printer_name)
        elif ext == 'pdf':
            print_pdf_to_printer(filepath, printer_name)
        elif ext in {'doc', 'docx'}:
            print_word_to_printer(filepath, printer_name)
        elif ext in {'xls', 'xlsx'}:
            print_excel_to_printer(filepath, printer_name)
        elif ext == 'txt':
            win32print.SetDefaultPrinter(printer_name)
            win32api.ShellExecute(0, "print", filepath, None, ".", 0)
        else:
            return "不支持的文件格式", 400

        winsound.MessageBeep()
        return redirect(url_for('index'))

    except Exception as e:
        return f"打印失败：{e}", 500

@app.route('/scan', methods=['GET'])
def scan():
    fmt = request.args.get("format", "pdf")
    ext = "pdf" if fmt == "pdf" else "png"
    dpi = request.args.get("dpi", "300")
    bitdepth = request.args.get("bitdepth", "color")  # color, gray, bw

    filename = f"scan_{datetime.now().strftime('%H%M%S')}.{ext}"
    output_file = os.path.join(SCAN_FOLDER, filename)

    naps2_path = r'"C:\Program Files\NAPS2\NAPS2.Console.exe"'
    cmd = f'{naps2_path} scan --output "{output_file}" --dpi {dpi} --bitdepth {bitdepth} --verbose'

    print("执行命令：", cmd)

    try:
        subprocess.run(cmd, shell=True, check=True)
    except subprocess.CalledProcessError as e:
        return f"扫描失败：{e}"

    if os.path.exists(output_file):
        return redirect(f"/preview/{filename}")
    else:
        return "扫描失败：文件未生成，请检查扫描仪或权限。"

@app.route('/preview/<filename>')
def preview(filename):
    return render_template('preview.html', filename=filename)

@app.route('/files/<filename>')
def serve_file(filename):
    filepath = os.path.join(SCAN_FOLDER, filename)
    if os.path.exists(filepath):
        return send_file(filepath)
    else:
        abort(404)


@app.route('/download/<filename>')
def download_and_delete(filename):
    filepath = os.path.join(SCAN_FOLDER, filename)
    if not os.path.exists(filepath):
        abort(404)

    with open(filepath, 'rb') as f:
        file_data = io.BytesIO(f.read())  # 把文件内容读进内存
        file_data.seek(0)  # 重置读指针

    @after_this_request
    def remove_file(response):
        try:
            os.remove(filepath)
            print(f"已删除扫描文件: {filepath}")
        except Exception as e:
            print(f"删除失败: {e}")
        return response

    return send_file(
        file_data,
        as_attachment=True,
        download_name=filename,
        mimetype='application/octet-stream'
    )


@app.route('/uploads/<filename>')
def view_upload(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

@app.route('/delete/<filename>')
def delete_file(filename):
    try:
        os.remove(os.path.join(UPLOAD_FOLDER, filename))
        return redirect(url_for('index'))
    except:
        return '删除失败', 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
