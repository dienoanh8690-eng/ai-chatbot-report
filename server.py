from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
import json
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from xhtml2pdf import pisa
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = "gemini-1.5-flash"  # Hoặc gemini-3.5-flash
GEMINI_API_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

# === THƯ MỤC ===
UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
KNOWLEDGE_FOLDER = "kho_kien_thuc"

for folder in [UPLOAD_FOLDER, RESULT_FOLDER, KNOWLEDGE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

INDEX_FILE = os.path.join(KNOWLEDGE_FOLDER, "danh_sach.json")
if not os.path.exists(INDEX_FILE):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump({"tai_lieu": [], "ket_qua": []}, f, ensure_ascii=False, indent=2)


# -------------------- ĐỌC FILE --------------------
def doc_file(duong_dan, dinh_dang):
    noi_dung = ""
    try:
        if dinh_dang == "docx":
            doc = Document(duong_dan)
            noi_dung = "\n".join([p.text for p in doc.paragraphs])
        elif dinh_dang == "xlsx":
            wb = load_workbook(duong_dan, data_only=True)
            ws = wb.active
            for hang in ws.iter_rows(values_only=True):
                noi_dung += " | ".join(str(c) if c else "" for c in hang) + "\n"
        elif dinh_dang in ["txt", "md"]:
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
        elif dinh_dang == "pdf":
            noi_dung = "[File PDF - nội dung đã lưu tham khảo]"
    except Exception as e:
        noi_dung = f"[Lỗi đọc: {str(e)}]"
    return noi_dung


# -------------------- LƯU LỊCH SỬ --------------------
def luu_vao_kho(loai, ten_file, mo_ta):
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    data[loai].append({
        "ten": ten_file,
        "mo_ta": mo_ta,
        "ngay": datetime.now().strftime("%d/%m/%Y %H:%M")
    })
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# -------------------- TẠO WORD --------------------
def tao_word(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    doc = Document()

    p = doc.add_heading("BÁO CÁO XỬ LÝ DỮ LIỆU THIẾT BỊ", 0)
    for run in p.runs:
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    p = doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    p.runs[0].font.name = "Arial"
    p._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    doc.add_paragraph("-" * 60)

    for dong in noi_dung.split("\n"):
        if dong.strip():
            p = doc.add_paragraph(dong)
            p.runs[0].font.name = "Arial"
            p._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    doc.save(duong_dan)
    return ten


# -------------------- TẠO EXCEL --------------------
def tao_excel(noi_dung=""):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    wb = Workbook()
    ws = wb.active
    ws.title = "DỮ LIỆU ĐÃ XỬ LÝ"

    in_dam = Font(bold=True, size=11, name="Arial")
    vien = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    can_giua = Alignment(horizontal='center', vertical='center')

    ws.merge_cells("A1:I1")
    ws["A1"] = "BÁO CÁO DỮ LIỆU THIẾT BỊ HỆ THỐNG"
    ws["A1"].font = Font(bold=True, size=14, color="0F4C81", name="Arial")
    ws["A1"].alignment = can_giua

    ws.merge_cells("A2:I2")
    ws["A2"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].alignment = can_giua

    cot = ["STT", "Mã thiết bị", "Tên thiết bị", "Quy cách", "Đơn vị", "Số lượng", "Đơn giá", "Thành tiền", "Ghi chú"]
    for c, ten_cot in enumerate(cot, 1):
        cell = ws.cell(row=4, column=c, value=ten_cot)
        cell.font = in_dam
        cell.alignment = can_giua
        cell.border = vien
        cell.fill = PatternFill("solid", fgColor="E6F2FF")

    hang = 5
    for dong in noi_dung.split("\n"):
        if dong.strip() and not dong.strip().startswith(("#", "---", "==")):
            ws.merge_cells(start_row=hang, start_column=1, end_row=hang, end_column=9)
            cell = ws.cell(row=hang, column=1, value=dong.strip())
            cell.alignment = Alignment(horizontal='left', vertical='center')
            hang += 1

    rong = [6, 12, 25, 20, 10, 10, 14, 14, 20]
    for c, w in enumerate(rong, 1):
        ws.column_dimensions[chr(64 + c)].width = w

    wb.save(duong_dan)
    return ten


# -------------------- TẠO PDF --------------------
def tao_pdf(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    duong_dan = os.path.join(RESULT_FOLDER, ten)

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; padding: 40px; line-height: 1.8; font-size: 14px; }}
            h1 {{ text-align: center; color: #0F4C81; border-bottom: 2px solid #0F4C81; padding-bottom: 10px; }}
            .ngay {{ text-align: right; color: #666; margin-bottom: 30px; }}
            hr {{ border: 1px solid #ccc; margin: 20px 0; }}
            pre {{ white-space: pre-wrap; font-family: inherit; }}
        </style>
    </head>
    <body>
        <h1>BÁO CÁO XỬ LÝ DỮ LIỆU THIẾT BỊ</h1>
        <p class="ngay">Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <hr>
        <pre>{noi_dung}</pre>
    </body>
    </html>
    """
    with open(duong_dan, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten


# -------------------- GỌI AI --------------------
def goi_ai(noi_dung, file_content=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY trên Render. Vào Environment Variables thêm khóa API."

    prompt = f"""Bạn là chuyên gia xử lý dữ liệu và lập báo cáo cho nhà máy thủy điện.

Yêu cầu: {noi_dung}

Nội dung file:
{file_content if file_content else '(Không có file)'}

Trả lời bằng tiếng Việt, rõ ràng, có cấu trúc."""

    try:
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        res = requests.post(GEMINI_API_URL, json=payload, timeout=60)

        if res.status_code != 200:
            return f"❌ API trả mã lỗi {res.status_code}: {res.text[:250]}"

        try:
            data = res.json()
        except Exception as e:
            return f"❌ API không trả JSON: {res.text[:250]}"

        if "candidates" not in data or not data["candidates"]:
            return f"❌ Không có kết quả từ AI: {json.dumps(data, ensure_ascii=False)}"

        return data["candidates"][0]["content"]["parts"][0]["text"]

    except requests.exceptions.Timeout:
        return "⏳ Hết thời gian chờ. Thử lại sau."
    except Exception as e:
        return f"❌ Lỗi kết nối: {str(e)}"


# ==================== ROUTE ====================
@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Không có file"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Chưa chọn file"}), 400

    ext = f.filename.rsplit(".", 1)[-1].lower()
    if ext not in ["docx", "xlsx", "txt", "pdf"]:
        return jsonify({"error": "Chỉ hỗ trợ .docx .xlsx .txt .pdf"}), 400

    ten_moi = f"{uuid.uuid4().hex[:10]}.{ext}"
    duong_dan = os.path.join(UPLOAD_FOLDER, ten_moi)
    f.save(duong_dan)

    noi_dung = doc_file(duong_dan, ext)
    luu_vao_kho("tai_lieu", ten_moi, f.filename)

    return jsonify({"status": "ok", "name": f.filename, "content": noi_dung[:3000]})


@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.json or {}
    cau_hoi = data.get("message", "").strip()
    file_content = data.get("file_content", "")

    if not cau_hoi and not file_content:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải file lên!"})

    tra_loi = goi_ai(cau_hoi, file_content)

    word = excel = pdf = ""
    if "❌" not in tra_loi and "⚠️" not in tra_loi and "⏳" not in tra_loi:
        word = tao_word(tra_loi)
        excel = tao_excel(tra_loi)
        pdf = tao_pdf(tra_loi)
        luu_vao_kho("ket_qua", word, cau_hoi[:100])

    return jsonify({
        "reply": tra_loi,
        "word": f"/download/{word}" if word else "",
        "excel": f"/download/{excel}" if excel else "",
        "pdf": f"/download/{pdf}" if pdf else ""
    })


@app.route("/download/<ten_file>")
def download(ten_file):
    for folder in [RESULT_FOLDER, UPLOAD_FOLDER]:
        path = os.path.join(folder, ten_file)
        if os.path.exists(path):
            return send_file(path, as_attachment=True)
    return "Không tìm thấy file", 404


@app.route("/")
def trang_chu():
    return """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>All Thủy Điện — Xử Lý Dữ Liệu Thông Minh</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #165DFF;
            --primary-dark: #0E42D2;
            --success: #00B42A;
            --warning: #FF7D00;
            --danger: #F53F3F;
            --bg: #F2F3F5;
            --card: #FFFFFF;
            --text-1: #1D2129;
            --text-2: #4E5969;
            --text-3: #86909C;
            --border: #E5E6EB;
            --shadow: 0 4px 24px rgba(0,0,0,0.08);
            --shadow-sm: 0 2px 8px rgba(0,0,0,0.05);
            --radius: 12px;
            --radius-lg: 16px;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { height: 100%; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background: linear-gradient(135deg, #E8F3FF 0%, #F2F3F5 100%);
            color: var(--text-1);
            line-height: 1.6;
            padding: 0;
            min-height: 100vh;
        }
        .page-wrapper {
            max-width: 720px;
            margin: 0 auto;
            padding: 32px 20px 48px;
        }
        .header { text-align: center; margin-bottom: 36px; }
        .logo-wrap {
            display: inline-flex; align-items: center; justify-content: center;
            width: 56px; height: 56px; border-radius: 16px;
            background: linear-gradient(135deg, #165DFF 0%, #4080FF 100%);
            color: white; font-size: 28px; font-weight: 700; margin-bottom: 12px;
            box-shadow: 0 8px 20px rgba(22, 93, 255, 0.25);
        }
        .header h1 {
            font-size: 26px; font-weight: 700; color: var(--text-1); margin-bottom: 4px;
        }
        .header p { font-size: 15px; color: var(--text-2); }
        .card {
            background: var(--card); border-radius: var(--radius-lg);
            box-shadow: var(--shadow); padding: 28px; margin-bottom: 20px;
            animation: cardIn .4s ease-out;
        }
        @keyframes cardIn {
            from { opacity: 0; transform: translateY(12px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .upload-zone {
            border: 2px dashed #C9CDD4; border-radius: var(--radius); padding: 36px 20px;
            text-align: center; cursor: pointer; transition: all 0.3s ease;
            background: #FAFAFA; margin-bottom: 16px;
        }
        .upload-zone:hover {
            border-color: var(--primary); background: #F0F5FF; transform: scale(1.01);
        }
        .upload-zone.active {
            border-color: var(--success); background: #E8FFEA; border-style: solid;
        }
        .upload-zone .icon { font-size: 32px; margin-bottom: 8px; }
        .upload-zone .title { font-weight: 600; color: var(--text-1); margin-bottom: 4px; }
        .upload-zone .sub { font-size: 13px; color: var(--text-3); }
        .file-info {
            display: flex; align-items: center; gap: 10px; padding: 12px 16px;
            background: #E8FFEA; border-radius: 8px; margin-bottom: 20px;
            display: none; border-left: 3px solid var(--success);
        }
        .file-info.show { display: flex; }
        .file-info .name { font-weight: 500; color: var(--success); }
        textarea {
            width: 100%; min-height: 110px; padding: 16px; border: 1px solid var(--border);
            border-radius: var(--radius); font-size: 15px; font-family: inherit;
            resize: vertical; margin-bottom: 16px; transition: border 0.2s, box-shadow 0.2s;
        }
        textarea:focus {
            outline: none; border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(22, 93, 255, 0.1);
        }
        .btn {
            width: 100%; padding: 14px 24px; border: none; border-radius: var(--radius);
            font-size: 16px; font-weight: 600; cursor: pointer; font-family: inherit;
            transition: all 0.2s ease;
        }
        .btn-primary {
            background: linear-gradient(90deg, var(--primary) 0%, #4080FF 100%);
            color: white; box-shadow: 0 4px 12px rgba(22, 93, 255, 0.25);
        }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(22, 93, 255, 0.35); }
        .btn-primary:disabled {
            background: #C9CDD4; cursor: not-allowed; transform: none; box-shadow: none;
        }
        .result-section { margin-top: 24px; display: none; }
        .result-section.show { display: block; animation: fadeIn 0.4s ease; }
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        .section-title {
            font-weight: 600; font-size: 15px; color: var(--text-1); margin-bottom: 12px;
            display: flex; align-items: center; gap: 8px;
        }
        .result-box {
            background: #F7F8FA; border-radius: var(--radius); padding: 20px;
            border-left: 4px solid var(--primary); white-space: pre-wrap;
            line-height: 1.8; font-size: 14px; max-height: 480px; overflow-y: auto;
            margin-bottom: 20px;
        }
        .result-box.error { border-left-color: var(--danger); background: #FFECEC; }
        .download-wrap {
            background: #F0F5FF; border-radius: var(--radius); padding: 20px;
            border: 1px solid #D6E4FF; display: none;
        }
        .download-wrap.show { display: block; animation: fadeIn 0.4s ease; }
        .download-grid {
            display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 12px;
        }
        @media (max-width: 520px) {
            .download-grid { grid-template-columns: 1fr; }
        }
        .download-btn {
            display: flex; flex-direction: column; align-items: center; gap: 6px;
            padding: 14px 10px; border-radius: 10px; text-decoration: none; font-weight: 600;
            transition: all 0.2s ease;
        }
        .download-btn:hover { transform: translateY(-2px); }
        .dw-word { background: #E8F3FF; color: var(--primary); }
        .dw-word:hover { background: #D6E8FF; }
        .dw-excel { background: #E8FFEA; color: var(--success); }
        .dw-excel:hover { background: #D6FFDB; }
        .dw-pdf { background: #FFECEC; color: var(--danger); }
        .dw-pdf:hover { background: #FFDBDB; }
        .dw-icon { font-size: 20px; }
        .dw-label { font-size: 13px; }
        .hint { font-size: 12px; color: var(--text-3); margin-top: -10px; margin-bottom: 14px; }
    </style>
</head>
<body>
    <div class="page-wrapper">
        <div class="header">
            <div class="logo-wrap">⚡</div>
            <h1>All Thủy Điện</h1>
            <p>Xử lý dữ liệu & lập báo cáo thông minh</p>
        </div>

        <div class="card">
            <div class="upload-zone" id="khuTai" onclick="document.getElementById('chonFile').click()">
                <div class="icon">📎</div>
                <div class="title">Tải tệp tài liệu lên</div>
                <div class="sub">Hỗ trợ: .docx .xlsx .txt .pdf</div>
                <input type="file" id="chonFile" accept=".docx,.xlsx,.txt,.pdf" style="display:none;" onchange="xuLyFile(this)">
            </div>

            <div class="file-info" id="thongTinFile">
                <span>✅</span>
                <span class="name" id="tenFile"></span>
            </div>

            <textarea id="cauhoi" placeholder="Nhập yêu cầu: kiểm tra, sắp xếp A-Z, tính thành tiền, tổng cộng..."></textarea>
            <p class="hint">Để trống chỉ tải file → hệ thống tự chuẩn hóa & tóm tắt</p>

            <button class="btn btn-primary" id="nutGui" onclick="gui()">
                <span id="nutText">🚀 Bắt đầu xử lý</span>
            </button>

            <div class="result-section" id="phanKetQua">
                <div class="section-title">📋 Kết quả xử lý</div>
                <div class="result-box" id="ketQua"></div>
            </div>

            <div class="download-wrap" id="khuTaiVe">
                <div class="section-title">💾 Tải kết quả về máy</div>
                <div class="download-grid">
                    <a id="btnWord" href="#" class="download-btn dw-word" target="_blank">
                        <span class="dw-icon">📄</span>
                        <span class="dw-label">Tải Word</span>
                    </a>
                    <a id="btnExcel" href="#" class="download-btn dw-excel" target="_blank">
                        <span class="dw-icon">📊</span>
                        <span class="dw-label">Tải Excel</span>
                    </a>
                    <a id="btnPdf" href="#" class="download-btn dw-pdf" target="_blank">
                        <span class="dw-icon">📕</span>
                        <span class="dw-label">Tải PDF</span>
                    </a>
                </div>
            </div>
        </div>
    </div>

    <script>
        let fileContent = "";

        async function xuLyFile(input) {
            const file = input.files[0];
            if (!file) return;
            const khu = document.getElementById("khuTai");
            const thongTin = document.getElementById("thongTinFile");
            
            khu.classList.remove("active");
            khu.innerHTML = `<div class="icon">⏳</div><div class="title">Đang đọc tệp...</div>`;
            
            const formData = new FormData();
            formData.append("file", file);
            
            try {
                const res = await fetch("/api/upload", { method: "POST", body: formData });
                const data = await res.json();
                
                if (data.status === "ok") {
                    fileContent = data.content;
                    khu.classList.add("active");
                    khu.innerHTML = `<div class="icon">✅</div><div class="title">Sẵn sàng nhận dữ liệu</div>`;
                    thongTin.classList.add("show");
                    document.getElementById("tenFile").textContent = data.name;
                } else {
                    khu.innerHTML = `<div class="icon">❌</div><div class="title">${data.error || 'Lỗi tải tệp'}</div>`;
                }
            } catch (e) {
                khu.innerHTML = `<div class="icon">❌</div><div class="title">Lỗi kết nối</div>`;
            }
        }

        async function gui() {
            const cauhoi = document.getElementById("cauhoi").value.trim();
            const nut = document.getElementById("nutGui");
            const nutText = document.getElementById("nutText");
            const phanKetQua = document.getElementById("phanKetQua");
            const ketQua = document.getElementById("ketQua");
            const khuTaiVe = document.getElementById("khuTaiVe");

            if (!cauhoi && !fileContent) {
                alert("Vui lòng nhập yêu cầu hoặc tải tệp lên!");
                return;
            }

            nut.disabled = true;
            nutText.textContent = "⏳ Đang xử lý...";
            phanKetQua.classList.remove("show");
            khuTaiVe.classList.remove("show");

            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ message: cauhoi, file_content: fileContent })
                });

                const text = await res.text();
                let data;
                try {
                    data = JSON.parse(text);
                } catch (e) {
                    phanKetQua.classList.add("show");
                    ketQua.classList.add("error");
                    ketQua.textContent = `❌ Phản hồi không hợp lệ:\\n${text.substring(0, 300)}`;
                    return;
                }

                ketQua.classList.remove("error");
                ketQua.textContent = data.reply || "Không có phản hồi";
                phanKetQua.classList.add("show");

                if (data.word || data.excel || data.pdf) {
                    khuTaiVe.classList.add("show");
                    if (data.word) document.getElementById("btnWord").href = data.word;
                    if (data.excel) document.getElementById("btnExcel").href = data.excel;
                    if (data.pdf) document.getElementById("btnPdf").href = data.pdf;
                }
            } catch (e) {
                phanKetQua.classList.add("show");
                ketQua.classList.add("error");
                ketQua.textContent = `❌ Lỗi: ${e.message}`;
            } finally {
                nut.disabled = false;
                nutText.textContent = "🚀 Bắt đầu xử lý";
            }
        }
    </script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
