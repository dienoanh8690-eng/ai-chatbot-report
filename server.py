from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
import json
from datetime import datetime
from docx import Document
from openpyxl import Workbook, load_workbook
from xhtml2pdf import pisa
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH BIẾN MÔI TRƯỜNG ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
FB_VERIFY_TOKEN = os.environ.get("FB_VERIFY_TOKEN", "baocao_ai_2026")
# ====================================================================

# === THƯ MỤC KHO KIẾN THỨC ===
GOC = "kho_kien_thuc_thuy_dien"
THU_MUC = {
    "goc": GOC,
    "mau_bao_cao": f"{GOC}/mau_bao_cao",
    "tai_lieu": f"{GOC}/tai_lieu_tham_khao",
    "lich_su": f"{GOC}/lich_su_hoat_dong",
    "du_lieu_da_huong": f"{GOC}/du_lieu_da_huong",
    "file_tao_ra": "generated_files",
    "file_tai_len": "uploaded_files"
}

for path in THU_MUC.values():
    os.makedirs(path, exist_ok=True)

INDEX_FILE = f"{GOC}/index_kien_thuc.json"
if not os.path.exists(INDEX_FILE):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump({"mau_bao_cao": [], "tai_lieu": [], "du_lieu_da_huong": []}, f, ensure_ascii=False, indent=2)


# -------------------- QUẢN LÝ CHỈ MỤC KHO --------------------
def doc_index():
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def luu_index(data):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def them_vao_index(loai, ten_file, mo_ta, tu_khoa):
    index = doc_index()
    index[loai].append({
        "ten_file": ten_file,
        "mo_ta": mo_ta,
        "tu_khoa": tu_khoa,
        "ngay_them": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "duong_dan": f"{THU_MUC[loai]}/{ten_file}"
    })
    luu_index(index)


# -------------------- TÌM KIẾM KIẾN THỨC LIÊN QUAN --------------------
def tim_kien_thuc(yeu_cau, gioi_han=5):
    index = doc_index()
    ket_qua = []
    tu_dong = set(yeu_cau.lower().split())
    
    for loai in ["mau_bao_cao", "tai_lieu", "du_lieu_da_huong"]:
        for muc in index.get(loai, []):
            diem = 0
            for tk in muc.get("tu_khoa", []):
                if any(tu in tk.lower() for tu in tu_dong) or tk.lower() in yeu_cau.lower():
                    diem += 1
            if diem > 0:
                ket_qua.append({"diem": diem, **muc})
    
    ket_qua.sort(key=lambda x: x["diem"], reverse=True)
    return ket_qua[:gioi_han]


# -------------------- ĐỌC NỘI DUNG FILE --------------------
def doc_noi_dung_file(duong_dan, dinh_dang):
    noi_dung = ""
    try:
        if dinh_dang == "docx":
            doc = Document(duong_dan)
            noi_dung = "\n".join([p.text for p in doc.paragraphs])
        elif dinh_dang == "xlsx":
            wb = load_workbook(duong_dan, data_only=True)
            ws = wb.active
            noi_dung = "\n".join([" | ".join(str(c) if c else "" for c in hang) for hang in ws.iter_rows(values_only=True)])
        elif dinh_dang == "txt":
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
        elif dinh_dang == "md":
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
        elif dinh_dang == "pdf":
            noi_dung = "[File PDF — nội dung được lưu tham khảo]"
    except Exception as e:
        noi_dung = f"[Lỗi đọc file: {str(e)}]"
    return noi_dung


# -------------------- LƯU KẾT QUẢ VÀO KHO --------------------
def luu_vao_kho(noi_dung_goc, noi_dung_ai, yeu_cau, loai="du_lieu_da_huong"):
    ten_file = f"phan_tich_{uuid.uuid4().hex[:10]}.md"
    duong_dan = f"{THU_MUC[loai]}/{ten_file}"
    
    noi_dung_luu = f"""# PHÂN TÍCH & BÁO CÁO THỦY ĐIỆN
Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}
Yêu cầu: {yeu_cau}

---
## NỘI DUNG GỐC
{noi_dung_goc[:3000] if noi_dung_goc else '(Không có file đính kèm)'}

---
## PHÂN TÍCH & KẾT QUẢ
{noi_dung_ai}
"""
    with open(duong_dan, "w", encoding="utf-8") as f:
        f.write(noi_dung_luu)
    
    tu_khoa = [t for t in yeu_cau.replace(",", " ").replace(".", " ").split() if len(t) > 2]
    them_vao_index(loai, ten_file, yeu_cau, tu_khoa)
    return ten_file


# -------------------- TẠO FILE WORD --------------------
def tao_word(noi_dung):
    ten_file = f"{THU_MUC['file_tao_ra']}/bao_cao_{uuid.uuid4().hex[:8]}.docx"
    doc = Document()
    doc.add_heading("BÁO CÁO KỸ THUẬT - NHÀ MÁY THỦY ĐIỆN", 0)
    doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph("=" * 60)
    for doan in noi_dung.split("\n"):
        if doan.strip():
            if doan.strip().startswith(("##", "###")):
                doc.add_heading(doan.strip().lstrip("# "), level=2)
            else:
                doc.add_paragraph(doan)
    doc.save(ten_file)
    return ten_file


# -------------------- TẠO FILE EXCEL --------------------
def tao_excel(noi_dung=""):
    ten_file = f"{THU_MUC['file_tao_ra']}/bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Báo cáo Kỹ thuật"
    ws.append(["BÁO CÁO KỸ THUẬT - NHÀ MÁY THỦY ĐIỆN"])
    ws.append(["Ngày tạo", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws.append([])
    for dong in noi_dung.split("\n"):
        if dong.strip():
            ws.append([dong.strip()])
    wb.save(ten_file)
    return ten_file


# -------------------- TẠO FILE PDF --------------------
def tao_pdf(noi_dung):
    ten_file = f"{THU_MUC['file_tao_ra']}/bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: DejaVu Sans; padding: 30px; line-height: 1.8; font-size: 14px; }}
            h1 {{ text-align: center; color: #0f4c81; border-bottom: 2px solid #0f4c81; padding-bottom: 10px; }}
            .ngay {{ color: #555; text-align: right; margin-bottom: 20px; }}
            h2 {{ color: #1e68a8; margin-top: 25px; }}
            hr {{ border: 1px solid #ccc; margin: 20px 0; }}
        </style>
    </head>
    <body>
        <h1>BÁO CÁO KỸ THUẬT — NHÀ MÁY THỦY ĐIỆN</h1>
        <p class="ngay">Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <hr>
        <p>{noi_dung.replace(chr(10), '<br>')}</p>
    </body>
    </html>
    """
    with open(ten_file, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten_file


# -------------------- GỌI GOOGLE GEMINI + KHO KIẾN THỨC --------------------
def goi_ai(noi_dung_nguoi_dung, noi_dung_file="", kien_thuc_lien_quan=None):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY. Lấy miễn phí tại: aistudio.google.com/apikey"
    
    tham_khao = ""
    if kien_thuc_lien_quan and len(kien_thuc_lien_quan) > 0:
        tham_khao = "\n\n📚 TÀI LIỆU THAM KHẢO TỪ KHO KIẾN THỨC:\n"
        for idx, kt in enumerate(kien_thuc_lien_quan, 1):
            nd = doc_noi_dung_file(kt["duong_dan"], kt["duong_dan"].split(".")[-1])
            tham_khao += f"---\n[{idx}] {kt['mo_ta']}:\n{nd[:1200]}\n"
    
    prompt = f"""Bạn là CHUYÊN GIA TƯ VẤN & THIẾT KẾ NHÀ MÁY THỦY ĐIỆN cấp cao.
Kinh nghiệm: thiết kế, tính toán, phân tích kỹ thuật, đánh giá an toàn, tối ưu hóa công trình thủy điện.
Ngôn ngữ: tiếng Việt chính thống, chuyên nghiệp, chuẩn mực kỹ thuật.

{tham_khao}

YÊU CẦU KHÁCH HÀNG:
{noi_dung_nguoi_dung}

NỘI DUNG TÀI LIỆU ĐÍNH KÈM:
{noi_dung_file if noi_dung_file else '(Không có file đính kèm)'}

---
Yêu cầu trả lời:
- Dùng kiến thức chuyên ngành thủy điện + tài liệu trong kho để phân tích sâu
- Cấu trúc rõ ràng: Mục tiêu → Phân tích → Kết luận → Khuyến nghị
- Dùng thuật ngữ kỹ thuật chuẩn, chính xác, có cơ sở
- Đề cập các tiêu chuẩn, quy trình tính toán thực tế
- Trả lời chi tiết, đầy đủ, có giá trị tham khảo thực tế
"""
    
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={GEMINI_API_KEY}"
        
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "maxOutputTokens": 4096,
                "temperature": 0.2
            }
        }
        
        response = requests.post(url, json=payload, timeout=180)
        
        if response.status_code == 200:
            result = response.json()
            return result["candidates"][0]["content"]["parts"][0]["text"]
        else:
            return f"❌ Lỗi API {response.status_code}: {response.text[:300]}"
    except Exception as e:
        return f"❌ Lỗi kết nối: {str(e)}"


# -------------------- TẢI FILE LÊN & LƯU VÀO KHO --------------------
@app.route("/api/upload", methods=["POST"])
def tai_file_len():
    if "file" not in request.files:
        return jsonify({"error": "Không có file nào được gửi"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Chưa chọn file"}), 400
    
    dinh_dang = file.filename.rsplit(".", 1)[-1].lower()
    ho_tro = ["docx", "xlsx", "txt", "pdf", "md"]
    
    if dinh_dang not in ho_tro:
        return jsonify({"error": f"Định dạng hỗ trợ: {', '.join(ho_tro)}"}), 400
    
    ten_file_moi = f"{uuid.uuid4().hex[:12]}.{dinh_dang}"
    duong_dan_luu = os.path.join(THU_MUC["tai_lieu"], ten_file_moi)
    file.save(duong_dan_luu)
    
    noi_dung = doc_noi_dung_file(duong_dan_luu, dinh_dang)
    
    # Thêm vào chỉ mục kho kiến thức
    tu_khoa = [t for t in file.filename.replace("_", " ").replace(".", " ").split() if len(t) > 2]
    them_vao_index("tai_lieu", ten_file_moi, f"Tải lên: {file.filename}", tu_khoa)
    
    return jsonify({
        "status": "ok",
        "filename": ten_file_moi,
        "preview": noi_dung[:600] + ("..." if len(noi_dung) > 600 else ""),
        "content": noi_dung,
        "luu_vao_kho": True
    })


# -------------------- LẤY DANH SÁCH KHO --------------------
@app.route("/api/kho", methods=["GET"])
def xem_kho():
    return jsonify(doc_index())


# -------------------- CHAT CHÍNH --------------------
@app.route("/api/chat", methods=["POST"])
def web_chat():
    data = request.json
    noi_dung = data.get("message", "").strip()
    noi_dung_file = data.get("file_content", "")
    
    if not noi_dung and not noi_dung_file:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải tài liệu lên!"})
    
    # Bước 1: Tìm kiến thức liên quan trong kho
    kien_thuc = tim_kien_thuc(noi_dung)
    
    # Bước 2: Gọi AI với kiến thức đã tìm được
    phan_hoi = goi_ai(noi_dung, noi_dung_file, kien_thuc)
    
    # Bước 3: Lưu toàn bộ hội thoại vào kho kiến thức
    luu_vao_kho(noi_dung_file, phan_hoi, noi_dung)
    
    # Bước 4: Tạo file kết quả
    word_link = excel_link = pdf_link = ""
    if "❌" not in phan_hoi and "⚠️" not in phan_hoi:
        wf = tao_word(phan_hoi)
        xf = tao_excel(phan_hoi)
        pf = tao_pdf(phan_hoi)
        word_link = f"/download/{os.path.basename(wf)}"
        excel_link = f"/download/{os.path.basename(xf)}"
        pdf_link = f"/download/{os.path.basename(pf)}"
    
    return jsonify({
        "reply": phan_hoi,
        "word": word_link,
        "excel": excel_link,
        "pdf": pdf_link,
        "tham_khao_da_tim": len(kien_thuc)
    })


# -------------------- TẢI FILE VỀ --------------------
@app.route("/download/<ten_file>")
def tai_file_ve(ten_file):
    for thu_muc in [THU_MUC["file_tao_ra"], THU_MUC["lich_su_hoat_dong"], THU_MUC["du_lieu_da_huong"]]:
        duong_dan = os.path.join(thu_muc, ten_file)
        if os.path.exists(duong_dan):
            return send_file(duong_dan, as_attachment=True)
    return "File không tồn tại", 404


# -------------------- TRANG CHỦ --------------------
@app.route("/")
def trang_chu():
    return """
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Chuyên gia Thủy Điện AI</title>
        <style>
            * { box-sizing: border-box; margin: 0; padding: 0; font-family: Arial, sans-serif; }
            body { max-width: 800px; margin: 30px auto; padding: 0 20px; background: #f0f7ff; }
            h1 { text-align: center; color: #0f4c81; margin-bottom: 5px; }
            .mota { text-align: center; color: #4a6a8a; margin-bottom: 20px; font-size: 15px; }
            .kho_info { background: #e6f2ff; padding: 10px 15px; border-radius: 8px; margin-bottom: 20px; border-left: 4px solid #0f4c81; }
            .khung { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 3px 12px rgba(15,76,129,0.1); }
            .tai_file { border: 2px dashed #94b8d9; padding: 25px; text-align: center; border-radius: 10px; margin-bottom: 18px; cursor: pointer; transition: 0.3s; }
            .tai_file:hover, .tai_file.dang_chon { border-color: #0f4c81; background: #e6f2ff; }
            textarea { width: 100%; height: 120px; padding: 14px; border: 1px solid #b3d1e8; border-radius: 8px; font-size: 15px; resize: vertical; margin-bottom: 15px; }
            button { background: #0f4c81; color: white; border: none; padding: 13px 32px; border-radius: 8px; font-size: 16px; cursor: pointer; font-weight: bold; }
            button:hover { background: #0d3c68; }
            button:disabled { background: #94b8d9; cursor: not-allowed; }
            .ketqua { margin-top: 25px; padding: 20px; background: #f0f9ff; border-radius: 8px; border-left: 4px solid #0f4c81; white-space: pre-wrap; line-height: 1.8; }
            .tai { margin-top: 15px; padding-top: 15px; border-top: 1px solid #cce0f0; }
            .tai a { display: inline-block; margin: 5px 15px 5px 0; color: #0f4c81; text-decoration: none; font-weight: bold; }
            .tai a:hover { text-decoration: underline; }
            .an { display: none; }
            .thong_bao { margin-top: 10px; font-size: 14px; color: #0f4c81; }
        </style>
    </head>
    <body>
        <h1>⚡ Chuyên gia Tư vấn & Thiết kế Thủy Điện AI</h1>
        <p class="mota">Kho kiến thức tích lũy — Trợ lý chuyên sâu lĩnh vực thủy điện</p>
        
        <div class="kho_info">
            📂 <strong>Kho kiến thức đang hoạt động:</strong> Tài liệu & kết quả sẽ được lưu và học hỏi liên tục
        </div>
        
        <div class="khung">
            <div class="tai_file" id="khuTai" onclick="document.getElementById('chonFile').click()">
                📎 <strong>Tải tài liệu kỹ thuật lên</strong><br>
                <span style="color:#666; font-size:13px;">Bản vẽ, tính toán, tiêu chuẩn, báo cáo (.docx .xlsx .txt .pdf)</span>
                <input type="file" id="chonFile" accept=".docx,.xlsx,.txt,.pdf,.md" style="display:none;" onchange="xuLyFile(this)">
            </div>
            <div id="tbFile" class="thong_bao an"></div>
            
            <textarea id="input" placeholder="Ví dụ: Phân tích thiết kế đập trọng lực bê tông cho nhà máy thủy điện công suất 50MW, lưu lượng trung bình 30m³/s..."></textarea>
            
            <br>
            <button id="nutgui" onclick="gui()">⚙️ Phân tích & Trả lời</button>
            
            <div id="ketqua" class="ketqua an"></div>
            <div id="tai" class="tai an"></div>
        </div>

        <script>
        let noiDungFile = "";

        async function xuLyFile(input) {
            const file = input.files[0];
            if (!file) return;
            
            const khu = document.getElementById("khuTai");
            const tb = document.getElementById("tbFile");
            
            khu.classList.add("dang_chon");
            khu.innerHTML = `✅ Đã chọn: <strong>${file.name}</strong><br><span style="color:#0f4c81">Đang lưu vào kho kiến thức...</span>`;
            tb.classList.add("an");
            
            const formData = new FormData();
            formData.append("file", file);
            
            try {
                const res = await fetch("/api/upload", { method: "POST", body: formData });
                const data = await res.json();
                
                if (data.status === "ok") {
                    noiDungFile = data.content;
                    khu.innerHTML = `✅ Đã lưu vào kho: <strong>${file.name}</strong><br><span style="color:#0f4c81">Nội dung sẵn sàng phân tích</span>`;
                } else {
                    khu.innerHTML = `❌ ${data.error}`;
                }
            } catch (e) {
                khu.innerHTML = `❌ Lỗi: ${e}`;
            }
        }

        async function gui() {
            const input = document.getElementById("input");
            const nut = document.getElementById("nutgui");
            const ketqua = document.getElementById("ketqua");
            const tai = document.getElementById("tai");
            
            const cau_hoi = input.value.trim();
            if (!cau_hoi && !noiDungFile) {
                alert("Vui lòng nhập yêu cầu kỹ thuật hoặc tải tài liệu lên!");
                return;
            }
            
            nut.disabled = true;
            nut.innerText = "Đang phân tích...";
            ketqua.style.display = "block";
            ketqua.innerText = "🔍 Tìm kiến thức trong kho...\n🤖 AI đang phân tích chuyên sâu...\n⏰ Vui lòng chờ 30–60 giây...";
            tai.style.display = "none";
            
            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ message: cau_hoi, file_content: noiDungFile })
                });
                const data = await res.json();
                
                ketqua.innerText = data.reply || "Không có phản hồi";
                
                if (data.tham_khao_da_tim > 0) {
                    ketqua.innerText = `📚 Đã tìm thấy ${data.tham_khao_da_tim} tài liệu tham khảo trong kho\n\n` + ketqua.innerText;
                }
                
                if (data.word || data.excel || data.pdf) {
                    tai.style.display = "block";
                    let html = "<strong>📂 Tải kết quả đầy đủ:</strong> ";
                    if (data.word) html += `<a href="${data.word}" target="_blank">📄 Word</a> `;
                    if (data.excel) html += `<a href="${data.excel}" target="_blank">📊 Excel</a> `;
                    if (data.pdf) html += `<a href="${data.pdf}" target="_blank">📕 PDF</a>`;
                    tai.innerHTML = html;
                }
            } catch (e) {
                ketqua.innerText = "❌ Lỗi kết nối: " + e;
            }
            
            nut.disabled = false;
            nut.innerText = "⚙️ Phân tích & Trả lời";
        }
        </script>
    </body>
    </html>
    """


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
