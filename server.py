from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
import json
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from xhtml2pdf import pisa
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# === THƯ MỤC LƯU DỮ LIỆU ===
UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
KNOWLEDGE_FOLDER = "kho_kien_thuc"

for folder in [UPLOAD_FOLDER, RESULT_FOLDER, KNOWLEDGE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

INDEX_FILE = os.path.join(KNOWLEDGE_FOLDER, "danh_sach.json")
if not os.path.exists(INDEX_FILE):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump({"tai_lieu": [], "ket_qua": []}, f, ensure_ascii=False, indent=2)


# -------------------- ĐỌC NỘI DUNG FILE --------------------
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


# -------------------- LƯU VÀO KHO KIẾN THỨC --------------------
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


# -------------------- TẠO FILE WORD — Đúng Font Tiếng Việt --------------------
def tao_word(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    doc = Document()
    
    # Tiêu đề
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


# -------------------- TẠO FILE EXCEL — Bảng Đầy Đủ + Công Thức --------------------
def tao_excel(noi_dung=""):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    wb = Workbook()
    ws = wb.active
    ws.title = "DỮ LIỆU ĐÃ XỬ LÝ"
    
    # Định dạng
    in_dam = Font(bold=True, size=11, name="Arial")
    vien = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    can_giua = Alignment(horizontal='center', vertical='center')
    can_trai = Alignment(horizontal='left', vertical='center')
    
    # Tiêu đề
    ws.merge_cells("A1:I1")
    ws["A1"] = "BÁO CÁO DỮ LIỆU THIẾT BỊ HỆ THỐNG"
    ws["A1"].font = Font(bold=True, size=14, color="0F4C81", name="Arial")
    ws["A1"].alignment = can_giua
    
    ws.merge_cells("A2:I2")
    ws["A2"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].alignment = can_giua
    
    # Tiêu đề cột
    cot = ["STT", "Mã thiết bị", "Tên thiết bị", "Quy cách", "Đơn vị", "Số lượng", "Đơn giá", "Thành tiền", "Ghi chú"]
    vi_tri = 4
    for c, ten_cot in enumerate(cot, 1):
        cell = ws.cell(row=vi_tri, column=c, value=ten_cot)
        cell.font = in_dam
        cell.alignment = can_giua
        cell.border = vien
    
    # Nội dung báo cáo -> tách thành các mục
    hang = vi_tri + 1
    for dong in noi_dung.split("\n"):
        if dong.strip() and not dong.strip().startswith(("#", "---", "==")):
            ws.merge_cells(start_row=hang, start_column=1, end_row=hang, end_column=9)
            cell = ws.cell(row=hang, column=1, value=dong.strip())
            cell.alignment = can_trai
            hang += 1
    
    # Dòng tổng cộng
    hang_tong = hang + 2
    ws.merge_cells(f"A{hang_tong}:G{hang_tong}")
    ws[f"A{hang_tong}"] = "TỔNG GIÁ TRỊ"
    ws[f"A{hang_tong}"].font = Font(bold=True, size=12, color="991B1B", name="Arial")
    
    # Độ rộng cột
    rong = [6, 12, 25, 20, 10, 10, 14, 14, 20]
    for c, w in enumerate(rong, 1):
        ws.column_dimensions[chr(64 + c)].width = w
    
    wb.save(duong_dan)
    return ten


# -------------------- TẠO FILE PDF — Font Tiếng Việt Đúng --------------------
def tao_pdf(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @font-face {{
                font-family: 'DejaVu Sans';
                src: url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');
            }}
            body {{ 
                font-family: 'DejaVu Sans', Arial, sans-serif; 
                padding: 40px; 
                line-height: 1.8;
                font-size: 14px;
            }}
            h1 {{ 
                text-align: center; 
                color: #0F4C81; 
                border-bottom: 2px solid #0F4C81; 
                padding-bottom: 10px;
                margin-bottom: 20px;
            }}
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
        pisa.CreatePDF(html, dest=f, encoding="utf-8")
    return ten


# -------------------- GỌI AI — Yêu cầu rõ ràng hơn --------------------
def goi_ai(noi_dung, file_content=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY"
    
    prompt = f"""Bạn là chuyên gia xử lý dữ liệu và lập báo cáo cho nhà máy thủy điện.

Yêu cầu người dùng:
{noi_dung}

Nội dung file đính kèm:
{file_content if file_content else '(Không có file)'}

---
Thực hiện chính xác theo các bước sau:

1. **Đánh giá dữ liệu đầu vào**:
   - Kiểm tra cấu trúc dữ liệu: hệ thống, thiết bị, số lượng, đơn giá...
   - Ghi rõ những điểm đúng, những điểm cần bổ sung/sửa

2. **Chuẩn hóa dữ liệu**:
   - Sửa lỗi chính tả, thống nhất định dạng ngày tháng, đơn vị tính
   - Sắp xếp danh sách thiết bị theo thứ tự A-Z theo tên hoặc mã thiết bị

3. **Tính toán**:
   - Thành tiền = Số lượng × Đơn giá
   - Tính tổng giá trị toàn bộ danh sách
   - Liệt kê rõ các thiết bị thiếu số lượng, thiếu đơn giá

4. **Trình bày báo cáo**:
   - Viết bằng tiếng Việt chuẩn, không lỗi font, rõ ràng
   - Cấu trúc: Mục đích → Đánh giá → Dữ liệu đã sắp xếp & tính toán → Kết luận
   - Trình bày chi tiết, đầy đủ, dễ đọc trên màn hình
"""
    
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={GEMINI_API_KEY}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"maxOutputTokens": 4096, "temperature": 0.2}
        }
        res = requests.post(url, json=payload, timeout=300)
        if res.status_code == 200:
            return res.json()["candidates"][0]["content"]["parts"][0]["text"]
        return f"❌ Lỗi API {res.status_code}: {res.text[:200]}"
    except Exception as e:
        return f"❌ Lỗi kết nối: {str(e)}"


# ==================== ROUTE ====================

@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Không có file"}), 400
    f = request.files["file"]
    if f.filename == "":
        return jsonify({"error": "Chưa chọn file"}), 400
    
    ext = f.filename.rsplit(".", 1)[-1].lower()
    if ext not in ["docx", "xlsx", "txt", "pdf"]:
        return jsonify({"error": "Chỉ hỗ trợ: .docx .xlsx .txt .pdf"}), 400
    
    ten_moi = f"{uuid.uuid4().hex[:10]}.{ext}"
    duong_dan = os.path.join(UPLOAD_FOLDER, ten_moi)
    f.save(duong_dan)
    
    noi_dung = doc_file(duong_dan, ext)
    luu_vao_kho("tai_lieu", ten_moi, f.filename)
    
    return jsonify({"status": "ok", "name": f.filename, "content": noi_dung[:3000]})


@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.json
    cau_hoi = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    
    if not cau_hoi and not file_content:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải file lên!"})
    
    tra_loi = goi_ai(cau_hoi, file_content)
    
    word = excel = pdf = ""
    if "❌" not in tra_loi and "⚠️" not in tra_loi:
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
    <title>All thủy điện</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: Arial, sans-serif; }
        body { max-width: 850px; margin: 30px auto; padding: 0 20px; background: #f0f7ff; }
        h1 { text-align: center; color: #0f4c81; margin-bottom: 30px; }
        .box { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        
        .upload-area { border: 2px dashed #94b8d9; padding: 25px; text-align: center; border-radius: 10px; cursor: pointer; margin-bottom: 15px; transition: 0.3s; }
        .upload-area:hover { border-color: #0f4c81; background: #e6f2ff; }
        .upload-area.active { border-color: #22c55e; background: #f0fdf4; }
        .file-info { margin: 10px 0 20px; padding: 10px 15px; background: #e6ffed; border-radius: 6px; display: none; color: #166534; font-weight: bold; }
        
        textarea { width: 100%; height: 110px; padding: 14px; border: 1px solid #b3d1e8; border-radius: 8px; font-size: 15px; margin-bottom: 15px; resize: vertical; }
        button { background: #0f4c81; color: white; border: none; padding: 13px 30px; border-radius: 8px; font-size: 16px; cursor: pointer; width: 100%; font-weight: bold; }
        button:hover { background: #0d3c68; }
        button:disabled { background: #94b8d9; cursor: not-allowed; }
        
        .result-section { margin-top: 25px; display: none; }
        .result-label { font-weight: bold; color: #0f4c81; margin-bottom: 10px; font-size: 16px; }
        .result-box { padding: 20px; background: #f8fbff; border-radius: 8px; border-left: 4px solid #0f4c81; white-space: pre-wrap; line-height: 1.7; max-height: 500px; overflow-y: auto; margin-bottom: 20px; }
        
        .download-box { padding: 15px 20px; background: #f0f9ff; border-radius: 8px; border: 1px solid #cce0f0; display: none; }
        .download-label { font-weight: bold; color: #0f4c81; margin-bottom: 12px; }
        .download-buttons { display: flex; gap: 12px; flex-wrap: wrap; }
        .download-btn { padding: 10px 20px; border-radius: 6px; text-decoration: none; font-weight: bold; display: inline-flex; align-items: center; gap: 8px; }
        .word { background: #e6f2ff; color: #0f4c81; }
        .excel { background: #e6ffed; color: #166534; }
        .pdf { background: #ffe6e6; color: #991b1b; }
        .download-btn:hover { transform: translateY(-2px); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    </style>
</head>
<body>
    <h1>⚡ All thủy điện</h1>
    <div class="box">
        <div class="upload-area" id="khuTai" onclick="document.getElementById('chonFile').click()">
            <strong>📎 Tải file tài liệu lên</strong><br>
            <span style="color:#666; font-size:13px;">.docx .xlsx .txt .pdf</span>
            <input type="file" id="chonFile" accept=".docx,.xlsx,.txt,.pdf" style="display:none;" onchange="xuLyFile(this)">
        </div>
        <div class="file-info" id="thongTinFile">✅ Đã chọn: <span id="tenFile"></span></div>
        
        <textarea id="cauhoi" placeholder="Ví dụ: Sắp xếp danh sách thiết bị A-Z, tính thành tiền, tổng cộng..."></textarea>
        
        <button id="nutGui" onclick="gui()">Gửi & Phân tích</button>
        
        <div class="result-section" id="phanKetQua">
            <div class="result-label">📋 Kết quả xử lý:</div>
            <div class="result-box" id="ketQua"></div>
        </div>
        
        <div class="download-box" id="khuTaiVe">
            <div class="download-label">💾 Tải kết quả về máy:</div>
            <div class="download-buttons">
                <a id="btnWord" href="#" class="download-btn word" target="_blank">📄 Tải Word</a>
                <a id="btnExcel" href="#" class="download-btn excel" target="_blank">📊 Tải Excel</a>
                <a id="btnPdf" href="#" class="download-btn pdf" target="_blank">📕 Tải PDF</a>
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
            khu.classList.add("active");
            khu.innerHTML = "⏳ Đang đọc file...";
            const formData = new FormData();
            formData.append("file", file);
            try {
                const res = await fetch("/api/upload", { method: "POST", body: formData });
                const data = await res.json();
                if (data.status === "ok") {
                    fileContent = data.content;
                    khu.innerHTML = "✅ Sẵn sàng nhận file";
                    thongTin.style.display = "block";
                    document.getElementById("tenFile").textContent = data.name;
                } else {
                    khu.innerHTML = "❌ " + data.error;
                }
            } catch (e) {
                khu.innerHTML = "❌ Lỗi: " + e;
            }
        }

        async function gui() {
            const cauhoi = document.getElementById("cauhoi").value.trim();
            const nut = document.getElementById("nutGui");
            const phanKetQua = document.getElementById("phanKetQua");
            const ketQua = document.getElementById("ketQua");
            const khuTaiVe = document.getElementById("khuTaiVe");
            
            if (!cauhoi && !fileContent) {
                alert("Vui lòng nhập yêu cầu hoặc tải file lên!");
                return;
            }
            
            nut.disabled = true;
            nut.textContent = "⏳ Đang xử lý...";
            phanKetQua.style.display = "none";
            khuTaiVe.style.display = "none";
            
            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ message: cauhoi, file_content: fileContent })
                });
                const data = await res.json();
                ketQua.textContent = data.reply || "Không có phản hồi";
                phanKetQua.style.display = "block";
                if (data.word || data.excel || data.pdf) {
                    khuTaiVe.style.display = "block";
                    if (data.word) document.getElementById("btnWord").href = data.word;
                    if (data.excel) document.getElementById("btnExcel").href = data.excel;
                    if (data.pdf) document.getElementById("btnPdf").href = data.pdf;
                }
            } catch (e) {
                phanKetQua.style.display = "block";
                ketQua.textContent = "❌ Lỗi kết nối: " + e;
            }
            
            nut.disabled = false;
            nut.textContent = "Gửi & Phân tích";
        }
    </script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
