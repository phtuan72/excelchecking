import logging
import pandas as pd
import os
import unicodedata
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def safe_float(x):
    """Chuyển đổi giá trị về float, xử lý lỗi định dạng."""
    try:
        return float(str(x).replace(',', '').strip())
    except:
        return None


def format_number(value):
    """Định dạng số theo kiểu có dấu phẩy phân cách hàng nghìn."""
    try:
        return "{:,.0f}".format(float(value))
    except (ValueError, TypeError):
        return value  # Nếu không thể chuyển đổi, giữ nguyên giá trị gốc.


def normalize_text(text):
    """Chuẩn hóa tên: chuyển về chữ thường, loại bỏ khoảng trắng và ký tự đặc biệt."""
    if pd.isna(text) or text is None:
        return ""
    return unicodedata.normalize('NFKD', str(text)).encode('ascii', errors='ignore').decode('utf-8').strip().lower()


@app.route('/', methods=['GET', 'POST'])
def index():
    result_html = ""
    prev_file_kiemtra = ""
    prev_file_chuan = ""
    sheets_kt = []
    sheets_chuan = []
    selected_sheet_kiemtra = ""
    selected_sheet_chuan = ""
    columns_kt_list = []
    columns_chuan_list = []
    selected_columns_kt = []  # Lưu giá trị của các dropdown file cần kiểm tra
    selected_columns_chuan = []  # Lưu giá trị của các dropdown file chuẩn

    if request.method == 'POST':
        logging.debug(f"📥 Dữ liệu nhận từ browser: {request.form}")

        # Lấy file từ file input nếu có; nếu không có, dùng giá trị đã lưu trong hidden field
        file_kt_obj = request.files.get('file_kiemtra')
        file_chuan_obj = request.files.get('file_chuan')
        prev_file_kiemtra = request.form.get('prev_file_kiemtra')
        prev_file_chuan = request.form.get('prev_file_chuan')

        # Xử lý file cần kiểm tra
        if file_kt_obj and file_kt_obj.filename:
            path_kt = os.path.join(UPLOAD_FOLDER, "temp_file_kiemtra.xlsx")
            file_kt_obj.save(path_kt)
        elif prev_file_kiemtra and os.path.exists(prev_file_kiemtra):
            path_kt = prev_file_kiemtra
        else:
            return "File cần kiểm tra chưa có."

        # Xử lý file chuẩn
        if file_chuan_obj and file_chuan_obj.filename:
            path_chuan = os.path.join(UPLOAD_FOLDER, "temp_file_chuan.xlsx")
            file_chuan_obj.save(path_chuan)
        elif prev_file_chuan and os.path.exists(prev_file_chuan):
            path_chuan = prev_file_chuan
        else:
            return "File chuẩn chưa có."

        # Lưu lại đường dẫn để tái sử dụng (hidden field)
        prev_file_kiemtra = path_kt
        prev_file_chuan = path_chuan

        # Lấy sheet đã chọn (nếu có); nếu chưa có, dùng mặc định khi lần đầu load
        # (mặc định: "DC9-P3" cho file cần kiểm tra, "summary" cho file chuẩn)
        selected_sheet_kiemtra = request.form.get('sheet_kiemtra') or "DC9-P3"
        selected_sheet_chuan = request.form.get('sheet_chuan') or "summary"

        try:
            # Đọc file Excel theo sheet đã chọn
            df_kt = pd.read_excel(path_kt, sheet_name=selected_sheet_kiemtra)
            df_chuan = pd.read_excel(path_chuan, sheet_name=selected_sheet_chuan)
        except Exception as e:
            logging.error(f"❌ Lỗi đọc file Excel: {e}")
            return f"Lỗi đọc file: {e}"

        # Lấy danh sách sheet (để hiển thị lại cho người dùng)
        try:
            sheets_kt = pd.ExcelFile(path_kt).sheet_names
            sheets_chuan = pd.ExcelFile(path_chuan).sheet_names
        except Exception as e:
            sheets_kt = []
            sheets_chuan = []

        # Lấy danh sách cột của sheet được chọn
        columns_kt_list = list(df_kt.columns)
        columns_chuan_list = list(df_chuan.columns)

        # Lấy danh sách cặp cột đã chọn từ form (với tên "col_kt[]" và "col_chuan[]")
        selected_columns_kt = request.form.getlist("col_kt[]")
        selected_columns_chuan = request.form.getlist("col_chuan[]")
        selected_columns = []
        for col_kt, col_chuan in zip(selected_columns_kt, selected_columns_chuan):
            if col_kt and col_chuan:
                selected_columns.append((col_kt, col_chuan))
        logging.debug(f"🧐 Cột được chọn từ browser: {selected_columns}")

        if not selected_columns:
            logging.warning("⚠️ Người dùng chưa chọn cặp cột nào!")
            result_html = "Vui lòng chọn ít nhất một cặp cột để so sánh!"
        else:
            # So sánh dữ liệu
            df_chuan['normalized_name'] = df_chuan['FullName'].apply(normalize_text)
            logging.debug(f"📋 Danh sách tên trong file chuẩn sau chuẩn hóa: {df_chuan['normalized_name'].tolist()}")
            unique_error_entries = set()
            html = [
                "<table class='table table-bordered table-striped'>",
                "<thead><tr>",
                "<th>Dòng</th>",
                "<th>FullName</th>",
                "<th>Lỗi</th>",
                "<th>Giá trị kiểm tra</th>",
                "<th>Giá trị chuẩn</th>",
                "</tr></thead><tbody>"
            ]
            errors = 0
            for i, row in df_kt.iterrows():
                if normalize_text(row.get('FullName', '')) == "total":
                    continue
                normalized_name = normalize_text(row.get('FullName', ''))
                matched = df_chuan[df_chuan['normalized_name'] == normalized_name]
                if matched.empty:
                    key = (i + 2, row.get('FullName', '').title(), "Không tìm thấy tên")
                    if key not in unique_error_entries:
                        unique_error_entries.add(key)
                        html.append(
                            f"<tr class='table-danger'><td>{i+2}</td>"
                            f"<td>{row['FullName'].title()}</td>"
                            f"<td>Không tìm thấy tên</td>"
                            f"<td>{row['FullName']}</td><td>-</td></tr>"
                        )
                        errors += 1
                    continue

                mrow = matched.iloc[0]
                for col_kt, col_chuan in selected_columns:
                    val_kt = row.get(col_kt, '')
                    val_chuan = mrow.get(col_chuan, '')
                    formatted_kt = format_number(val_kt)
                    formatted_chuan = format_number(val_chuan)
                    logging.debug(f"🔍 So sánh dòng {i+2} | Cột kiểm tra `{col_kt}` ({formatted_kt}) vs "
                                  f"Cột chuẩn `{col_chuan}` ({formatted_chuan})")
                    if safe_float(val_kt) != safe_float(val_chuan):
                        key = (i + 2, row['FullName'].title(), col_kt, formatted_kt, formatted_chuan)
                        if key not in unique_error_entries:
                            unique_error_entries.add(key)
                            html.append(
                                f"<tr class='table-warning'><td>{i+2}</td>"
                                f"<td>{row['FullName'].title()}</td>"
                                f"<td>{col_kt}</td>"
                                f"<td>{formatted_kt}</td>"
                                f"<td>{formatted_chuan}</td></tr>"
                            )
                            errors += 1
            html.append("</tbody></table>")
            result_html = "\n".join(html)

    return render_template("index.html", result=result_html,
                           prev_file_kiemtra=prev_file_kiemtra,
                           prev_file_chuan=prev_file_chuan,
                           sheets_kiemtra=sheets_kt,
                           sheets_chuan=sheets_chuan,
                           selected_sheet_kiemtra=selected_sheet_kiemtra,
                           selected_sheet_chuan=selected_sheet_chuan,
                           columns_kt_list=columns_kt_list,
                           columns_chuan_list=columns_chuan_list,
                           selected_columns_kt=selected_columns_kt,
                           selected_columns_chuan=selected_columns_chuan)


if __name__ == '__main__':
    app.run(debug=True)
