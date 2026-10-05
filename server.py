import os
import re
import uuid
from datetime import datetime
import requests
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from docx import Document
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
CORS(app)

# ==================================================
# 🔑 GIỮ NGUYÊN KHÓA CỦA BẠN — ĐÃ CHẠY ĐƯỢC HÔM TRƯỚC
# ==================================================
GEMINI_API_KEY = os.environ.get(
    "GEMINI_API_KEY",
    "AQ.AB8RN6K1GYFuRSF0fSXz6gnjpQK9MWWT9ePrK967w2ycgMMYQ"
).strip()

GROQ_API_KEY = os.environ.get(
    "GROQ_API_KEY",
    "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"
).strip()

GEMINI_MODEL = "gemini-3.5-pro"  # ← Đúng phiên bản bạn đã dùng!
TIMEOUT = 30

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

MAC_DINH = "Trả lời bằng tiếng Việt rõ ràng, tự nhiên, dễ hiểu, chính xác."

# ==================================================
# GỌI GEMINI — THEO KIỂU ĐÃ CHẠY ĐƯỢC HÔM TRƯỚC
# ==================================================
def goi_gemini(prompt, he_thong=""):
    ten = "🔵 Gemini 3.5"
    
    # Thử lần lượt các endpoint tương thích với khóa AQ...
    endpoints = [
        f"https://generativelanguage.googleapis.com/v1alpha/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}",
        f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}",
        f"https://generativelanguage.googleapis.com/v1/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}",
    ]
    
    for url in endpoints:
        try:
            res = requests.post(
                url,
                headers={"Content-Type": "application/json"},
                json={
                    "system_instruction": {"parts": [{"text": he_thong or MAC_DINH}]},
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generation_config": {"temperature": 0.7, "max_output_tokens": 2048},
                },
                timeout=TIMEOUT,
            )
            
            if res.status_code == 200:
                data = res.json()
                cands = data.get("candidates", [])
                if cands:
                    parts = cands[0].get("content", {}).get("parts", [])
                    text = "".join(p.get("text", "") for p in parts).strip()
                    if text:
                        return ten, text, None
        except Exception as e:
            continue
    
    return ten, None, "❌ Tất cả đường dẫn đều không tương thích — kiểm tra lại khóa hoặc phiên bản mô hình"

# ==================================================
# GỌI GROQ — Dự phòng
# ==================================================
def goi_groq(prompt, he_thong=""):
    ten = "🟢 Groq/Llama"
    if not GROQ_API_KEY or not GROQ_API_KEY.startswith("gsk_"):
        return ten, None, "Chưa điền khóa Groq"
    
    url = "https://api.groq.com/openai/v1/chat/completions"
    try:
        res = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {GROQ_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": "llama-3.1-8b-instant",
                "messages": [
                    {"role": "system", "content": he_thong or MAC_DINH},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.7,
                "max_tokens": 2048,
            },
            timeout=TIMEOUT,
        )
        if res.status_code == 200:
            return ten, res.json()["choices"][0]["message"]["content"].strip(), None
    except Exception as e:
        pass
    return ten, None, "Không kết nối được"

# ==================================================
# TỰ ĐỘNG CHUYỂN ĐỔI — Gemini trước, Groq sau
# ==================================================
def goi_tu_dong(prompt, danh_sach, he_thong=""):
    loi_tong = []
    for ham in danh_sach:
        ten, kq, loi = ham(prompt, he_thong)
        if kq:
            return {"ok": True, "reply": f"✅ [{ten}]\n{kq}", "plain": kq}
        loi_tong.append(f"× {ten}: {loi}")
    return {"ok": False, "reply": "❌ Tất cả AI đều không trả lời:\n" + "\n".join(loi_tong), "plain": ""}

# ==================================================
# CHUYÊN MỤC
# ==================================================
CHATS = {
    "general": (
        "Bạn là trợ lý AI tổng hợp thông minh, hữu ích. Trả lời rõ ràng, dễ hiểu, bằng tiếng Việt.",
        [goi_gemini, goi_groq]
    ),
    "data": (
        "Bạn là chuyên gia phân tích dữ liệu và lập báo cáo. Tóm tắt số liệu, dùng bảng khi phù hợp, ngắn gọn.",
        [goi_gemini, goi_groq]
    ),
    "doc": (
        "Bạn là chuyên gia soạn thảo văn bản theo chuẩn Việt Nam. Viết trang trọng, đúng thể thức, rõ ràng.",
        [goi_gemini, goi_groq]
    ),
    "tender": (
        "Bạn là chuyên gia tư vấn đấu thầu theo pháp luật Việt Nam. Hướng dẫn chi tiết từng bước, dễ hiểu.",
        [goi_gemini, goi_groq]
    ),
    "equip": (
        "Bạn là chuyên gia quản lý thiết bị nhà máy thủy điện. Phân tích, đề xuất bảo trì, thay thế bằng tiếng Việt.",
        [goi_gemini, goi_groq]
    ),
}

# ==================================================
# XUẤT WORD & EXCEL
# ==================================================
def _sach(dong):
    return re.sub(r"[*#`]+", "", dong).strip()

def tao_word(noi_dung):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
        doc = Document()
        doc.add_heading("BÁO CÁO", 0)
        doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        for dong in noi_dung.split("\n"):
            d = _sach(dong)
            if d and not re.fullmatch(r"[|\-:\s]+", d):
                doc.add_paragraph(d)
        doc.save(os.path.join(RESULT_FOLDER, ten))
        return ten
    except Exception as e:
        print(f"Lỗi tạo Word: {e}", flush=True)
        return ""

def tao_excel(noi_dung):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "BÁO CÁO"
        ws["A1"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        hang = 3
        for dong in noi_dung.split("\n"):
            d = dong.strip()
            if not d: continue
            if d.startswith("|"):
                if re.fullmatch(r"[|\-:\s]+", d): continue
                for i, cell in enumerate(d.strip("|").split("|"), start=1):
                    ws.cell(row=hang, column=i, value=_sach(cell))
            else:
                ws.cell(row=hang, column=1, value=_sach(d))
            hang += 1
        wb.save(os.path.join(RESULT_FOLDER, ten))
        return ten
    except Exception as e:
        print(f"Lỗi tạo Excel: {e}", flush=True)
        return ""

def doc_google_sheet(sheet_url):
    try:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", sheet_url)
        if not m:
            return None, "❌ Link không đúng định dạng Google Sheets"
        csv_url = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv"
        res = requests.get(csv_url, timeout=15)
        if res.status_code == 200:
            res.encoding = "utf-8"
            return res.text, None
        return None, "❌ Chia sẻ Sheet → Bất kỳ ai có đường liên kết"
    except Exception as e:
        return None, f"❌ Lỗi: {e}"

# ==================================================
# API
# ==================================================
@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Không có tệp"}), 400
    tep = request.files["file"]
    ten_goc = tep.filename or "tep"
    ext = ten_goc.rsplit(".", 1)[-1].lower() if "." in ten_goc else ""
    if ext not in ("xlsx", "csv", "txt"):
        return jsonify({"error": "Chỉ hỗ trợ .xlsx, .csv, .txt"}), 400
    duong_dan = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex[:10]}.{ext}")
    tep.save(duong_dan)
    noi_dung = ""
    try:
        if ext == "xlsx":
            wb = load_workbook(duong_dan, data_only=True, read_only=True)
            for ws in wb.worksheets:
                noi_dung += f"=== Trang tính: {ws.title} ===\n"
                for hang in ws.iter_rows(values_only=True):
                    noi_dung += " | ".join("" if c is None else str(c) for c in hang) + "\n"
                    if len(noi_dung) > 6000: break
                if len(noi_dung) > 6000: break
            wb.close()
        else:
            with open(duong_dan, "r", encoding="utf-8-sig", errors="ignore") as fh:
                noi_dung = fh.read()
    except Exception as e:
        noi_dung = f"(Không đọc được: {e})"
    return jsonify({"status": "ok", "name": ten_goc, "content": noi_dung[:6000]})

@app.route("/api/connect-sheet", methods=["POST"])
def connect_sheet():
    url = ((request.get_json(silent=True) or {}).get("url") or "").strip()
    if not url:
        return jsonify({"error": "Dán link Google Sheets nhé"}), 400
    noi_dung, loi = doc_google_sheet(url)
    if loi:
        return jsonify({"error": loi}), 400
    return jsonify({"status": "ok", "content": noi_dung[:6000]})

@app.route("/api/chat/<kind>", methods=["POST"])
def chat(kind):
    if kind not in CHATS:
        return jsonify({"reply": "❌ Chức năng không tồn tại"}), 404
    data = request.get_json(silent=True) or {}
    msg = (data.get("message") or "").strip()
    ctx = "\n---\n".join(x for x in [(data.get("file_content") or "")[:6000],
                                      (data.get("sheet_content") or "")[:6000]] if x.strip())
    if not msg and not ctx:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải dữ liệu!"})
    prompt = f"{msg or 'Hãy phân tích dữ liệu sau.'}\n\nDữ liệu:\n{ctx}" if ctx else msg
    he_thong, danh_sach = CHATS[kind]
    kq = goi_tu_dong(prompt, danh_sach, he_thong)
    out = {"reply": kq["reply"], "word": "", "excel": ""}
    if kind == "data" and kq["ok"]:
        w, x = tao_word(kq["plain"]), tao_excel(kq["plain"])
        out["word"] = f"/download/{w}" if w else ""
        out["excel"] = f"/download/{x}" if x else ""
    return jsonify(out)

@app.route("/download/<path:ten_file>")
def download(ten_file):
    return send_from_directory(RESULT_FOLDER, os.path.basename(ten_file), as_attachment=True)

@app.route("/healthz")
def healthz():
    return "ok"

# ==================================================
# GIAO DIỆN
# ==================================================
@app.route("/")
def trang_chu():
    return TRANG_CHU

TRANG_CHU = r'''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Thủy Điện — Hệ thống hỗ trợ</title>
<style>
:root{--g100:#f1f5f9;--g200:#e2e8f0;--g600:#475569;--g800:#1e293b}
*{margin:0;padding:0;box-sizing:border-box;font-family:Inter,"Segoe UI",Roboto,Arial,sans-serif}
body{background:linear-gradient(135deg,#f0f7ff,#faf5ff);min-height:100vh;padding:20px}
.header{text-align:center;margin-bottom:28px;padding-top:10px}
.header h1{font-size:26px;color:var(--g800);margin-bottom:6px}
.header p{color:var(--g600);font-size:14px}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:22px;max-width:2200px;margin:0 auto}
@media(max-width:1400px){.grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:768px){.grid{grid-template-columns:1fr}}
.card{background:#fff;border-radius:20px;padding:24px;box-shadow:0 4px 12px rgba(0,0,0,.06);display:flex;flex-direction:column;height:720px;border-top:4px solid var(--c)}
.c-gen{--c:#6366f1;--cl:#e0e7ff}.c-p1{--c:#2563eb;--cl:#dbeafe}
.c-p2{--c:#16a34a;--cl:#dcfce7}.c-p3{--c:#9333ea;--cl:#f3e8ff}
.c-p4{--c:#f59e0b;--cl:#fef3c7}.c-p5{--c:#ec4899;--cl:#fce7f3}
.card-head{display:flex;align-items:center;gap:10px;margin-bottom:18px;padding-bottom:14px;border-bottom:2px solid var(--cl);flex-wrap:wrap}
.card-icon{font-size:24px}.card-title{font-size:17px;font-weight:700;color:var(--c)}
.card-ai{font-size:11px;color:#94a3b8;margin-left:auto;background:var(--g100);padding:3px 8px;border-radius:12px}
.upload-zone{border:2px dashed var(--g200);border-radius:12px;padding:14px;text-align:center;cursor:pointer;margin-bottom:12px}
.upload-zone:hover,.upload-zone.drag{border-color:var(--c);background:var(--cl)}
.file-bar{display:flex;align-items:center;gap:8px;padding:10px 14px;border-radius:8px;margin-bottom:12px;font-size:13px;background:var(--cl)}
.file-bar button{margin-left:auto;background:none;border:none;font-size:18px;cursor:pointer;color:#dc2626}
.sheet-bar{display:flex;gap:8px;margin-bottom:12px}
.sheet-bar input{flex:1;padding:10px 14px;border:1px solid var(--g200);border-radius:8px;font-size:13px}
.sheet-bar button{padding:10px 16px;border:none;border-radius:8px;background:var(--c);color:#fff;cursor:pointer;font-weight:600}
.quick-btns{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px}
.q-btn{padding:10px 12px;border:1px solid var(--g200);border-radius:8px;background:#fff;cursor:pointer;font-size:12px;text-align:left}
.q-btn:hover{border-color:var(--c);background:var(--cl)}
.chat-area{flex:1;overflow-y:auto;padding:4px;margin-bottom:14px;min-height:200px}
.msg{margin-bottom:16px;max-width:98%}.msg.user{margin-left:auto}
.bubble{padding:14px 18px;border-radius:18px;line-height:1.6;font-size:14px;white-space:pre-wrap}
.msg.user .bubble{background:linear-gradient(135deg,#dbeafe,#e0e7ff);border-bottom-right-radius:8px}
.msg.ai .bubble{background:var(--g100);border-bottom-left-radius:8px}
.bubble.err{background:#fee2e2;border-left:3px solid #ef4444}
.dl-group{display:flex;gap:10px;margin-top:12px;flex-wrap:wrap}
.dl-btn{padding:8px 16px;border-radius:20px;text-decoration:none;font-size:13px;font-weight:600}
.dl-word{background:#dbeafe;color:#1d4ed8}.dl-excel{background:#dcfce7;color:#15803d}
.input-row{display:flex;gap:10px;align-items:flex-end}
textarea{flex:1;min-height:48px;max-height:120px;padding:12px 18px;border:1px solid var(--g200);border-radius:24px;font-size:14px;resize:none;outline:none}
textarea:focus{border-color:var(--c)}
.send-btn{width:44px;height:44px;border-radius:50%;border:none;color:#fff;background:var(--c);cursor:pointer;font-size:18px}
.send-btn:disabled{opacity:.5;cursor:not-allowed}
.hidden{display:none!important}
.contact-row{display:flex;align-items:flex-start;gap:10px;padding:8px 0;border-bottom:1px solid var(--g100);font-size:14px}
.contact-label{font-weight:600;color:var(--c);min-width:90px}
a{color:var(--c);text-decoration:none}
</style>
</head>
<body>
<div class="header">
<h1>⚡ All Thủy Điện — Hệ thống hỗ trợ</h1>
<p>Trò chuyện · Phân tích dữ liệu · Soạn thảo · Đấu thầu · Quản lý thiết bị — <span style="color:#16a34a;font-weight:bold">AI miễn phí</span></p>
</div>
<div class="grid" id="grid">
<div class="card c-p5" id="contactCard">
<div class="card-head"><span class="card-icon">📌</span><h3 class="card-title">Thông tin liên hệ</h3></div>
<div>
<div class="contact-row"><span>👤</span><span class="contact-label">Admin:</span><span>Chotvjp</span></div>
<div class="contact-row"><span>📞</span><span class="contact-label">ĐT:</span><a href="tel:0973020486">0973020486</a></div>
<div class="contact-row"><span>✉️</span><span class="contact-label">Email:</span><a href="mailto:hoangdien86ncc@gmail.com">hoangdien86ncc@gmail.com</a></div>
<div class="contact-row"><span>🏢</span><span class="contact-label">Đơn vị:</span><span>Công ty Cổ phần Thủy điện Nậm Chiến</span></div>
<div class="contact-row"><span>📍</span><span class="contact-label">Địa chỉ:</span><span>TK5 - Mường La - Sơn La</span></div>
<div class="contact-row"><span>🌐</span><span class="contact-label">Web:</span><a href="https://namchien.vn" target="_blank">namchien.vn</a></div>
</div>
<div style="margin-top:auto;padding-top:16px;text-align:center;font-size:12px;color:#9ca3af">
<p>© 2026 — Hệ thống hỗ trợ công việc nội bộ</p>
</div>
</div>
</div>
<script>
const CARDS = [
 {id:"general",cls:"c-gen",icon:"💬",title:"Trò chuyện chung",ai:"🔵 Gemini 3.5 → 🟢 Groq",
  hello:"👋 Xin chào! Tôi đã sẵn sàng. Hỏi tôi bất kỳ điều gì nhé!",ph:"Nhập câu hỏi...",
  quick:[["🔌 Thủy điện cơ bản","Giải thích khái niệm nhà máy thủy điện"],["📊 Hiệu suất & Tối ưu","Cách nâng cao hiệu suất làm việc"],["📋 Quy trình chung","Trình bày quy trình làm việc chuẩn"],["💡 Ý tưởng & Đề xuất","Đề xuất ý tưởng cải tiến"]]},
 {id:"data",cls:"c-p1",icon:"📊",title:"Xử lý dữ liệu & Tạo báo cáo",ai:"🔵 Gemini 3.5 → 🟢 Groq",file:true,sheet:true,
  hello:"👋 Tải tệp, dán link Sheets hoặc nhập yêu cầu — tôi phân tích và xuất Word/Excel nhé!",ph:"Nhập yêu cầu phân tích...",
  quick:[["📋 Tóm tắt dữ liệu","Tóm tắt số liệu chính"],["💰 Tính tổng hợp","Tính tổng, trung bình, xu hướng"],["📑 Báo cáo đầy đủ","Viết báo cáo có cấu trúc"],["📈 Nhận xét & Đề xuất","Đánh giá và đề xuất"]]},
 {id:"doc",cls:"c-p2",icon:"✍️",title:"Soạn thảo văn bản",ai:"🔵 Gemini 3.5 → 🟢 Groq",
  hello:"👋 Tôi soạn thảo văn bản chuẩn mực Việt Nam. Bạn cần viết gì?",ph:"Nội dung cần soạn thảo...",
  quick:[["📝 Công văn hành chính","Soạn thảo công văn gửi cấp trên"],["📄 Hợp đồng & Thỏa thuận","Soạn thảo hợp đồng mua bán dịch vụ"],["📈 Báo cáo công việc","Báo cáo tiến độ, kết quả thực hiện"],["📋 Thư mời & Biên bản","Thư mời họp, biên bản cuộc họp"]]},
 {id:"tender",cls:"c-p3",icon:"🏆",title:"Quy trình đấu thầu",ai:"🔵 Gemini 3.5 → 🟢 Groq",
  hello:"👋 Tôi hướng dẫn theo Luật Đấu thầu Việt Nam. Cần hỗ trợ bước nào?",ph:"Hỏi về quy trình đấu thầu...",
  quick:[["📋 Toàn bộ quy trình","Giải thích các bước từ A-Z"],["📑 Hồ sơ mời thầu","Danh mục tài liệu cần chuẩn bị"],["⚖️ Pháp lý & Lưu ý","Điều khoản pháp lý thường gặp"],["📄 Mẫu biểu thông dụng","Danh sách biểu mẫu cần có"]]},
 {id:"equip",cls:"c-p4",icon:"🔧",title:"Quản lý thiết bị",ai:"🔵 Gemini 3.5 → 🟢 Groq",file:true,
  hello:"👋 Tải danh sách thiết bị hoặc nhập yêu cầu — tôi phân tích nhé!",ph:"Nhập yêu cầu quản lý thiết bị...",
  quick:[["📊 Phân loại thiết bị","Nhóm thiết bị theo chức năng"],["🛠️ Kế hoạch bảo trì","Lập kế hoạch bảo trì định kỳ"],["⚠️ Đánh giá tình trạng","Phân tích rủi ro & đề xuất"],["🔄 Tuổi thọ & Thay thế","Tính tuổi thọ, đề xuất thay thế"]]}
];
const $ = id => document.getElementById(id);
const ST = {};
function cardHTML(c){
  return '<div class="card-head"><span class="card-icon">'+c.icon+'</span><h3 class="card-title">'+c.title+'</h3><span class="card-ai">'+c.ai+'</span></div>'
  +(c.file?'<div class="upload-zone" id="up-'+c.id+'"><p>📎 Nhấn chọn hoặc kéo thả tệp (.xlsx, .csv, .txt)</p></div><input type="file" id="fi-'+c.id+'" accept=".xlsx,.csv,.txt" class="hidden"><div class="file-bar hidden" id="fb-'+c.id+'"><span></span><button>✕</button></div>':"")
  +(c.sheet?'<div class="sheet-bar"><input type="text" id="su-'+c.id+'" placeholder="🔗 Dán link Google Sheets..."><button id="sb-'+c.id+'">Kết nối</button></div><div class="file-bar hidden" id="sh-'+c.id+'"><span>✅ Đã kết nối Sheets</span><button>✕</button></div>':"")
  +'<div class="quick-btns">'+c.quick.map((q,k)=>'<button class="q-btn" data-k="'+k+'">'+q[0]+'</button>').join("")+'</div>'
  +'<div class="chat-area" id="ch-'+c.id+'"></div>'
  +'<div class="input-row"><textarea id="in-'+c.id+'" placeholder="'+c.ph+'"></textarea><button class="send-btn" id="sd-'+c.id+'">➤</button></div>';
}
function addMsg(id, type, text, d){
  const m = document.createElement("div"); m.className = "msg "+type;
  const b = document.createElement("div"); b.className = "bubble"+(text.includes("❌")?" err":"");
  b.textContent = text; m.appendChild(b);
  if(d && (d.word || d.excel)){
    const g = document.createElement("div"); g.className = "dl-group";
    if(d.word) g.innerHTML += '<a class="dl-btn dl-word" href="'+d.word+'" download>📄 Tải Word</a>';
    if(d.excel) g.innerHTML += '<a class="dl-btn dl-excel" href="'+d.excel+'" download>📊 Tải Excel</a>';
    b.appendChild(g);
  }
  const c = $("ch-"+id); c.appendChild(m); c.scrollTop = c.scrollHeight;
}
async function call(path, opts){
  opts.headers = Object.assign({}, opts.headers || {});
  const r = await fetch(path, opts);
  return r.json();
}
async function send(c, preset){
  const i = $("in-"+c.id), s = ST[c.id];
  const m = (preset !== undefined ? preset : i.value).trim();
  if(!m && !s.file && !s.sheet) return;
  addMsg(c.id, "user", m || "Phân tích dữ liệu"); i.value = "";
  const btn = $("sd-"+c.id); btn.disabled = true; btn.textContent = "⏳";
  try{
    const d = await call("/api/chat/"+c.id, {method:"POST", headers:{"Content-Type":"application/json"},
      body: JSON.stringify({message:m, file_content:s.file, sheet_content:s.sheet})});
    addMsg(c.id, "ai", d.reply || "❌ Không có phản hồi", d);
  }catch(e){
    addMsg(c.id, "ai", "❌ Không kết nối được — nếu mới khởi động, vui lòng chờ 30-60 giây thử lại nhé!");
  }finally{
    btn.disabled = false; btn.textContent = "➤";
  }
}
async function upload(c, f){
  const fd = new FormData(); fd.append("file", f);
  try{
    const d = await call("/api/upload", {method:"POST", body:fd});
    if(d.status === "ok"){
      ST[c.id].file = d.content;
      $("fb-"+c.id).firstChild.textContent = "📎 "+d.name;
      $("fb-"+c.id).classList.remove("hidden");
    } else alert(d.error);
  }catch(e){ alert("Không tải được tệp"); }
}
function wire(c){
  $("in-"+c.id).addEventListener("keydown", e => { if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); send(c); } });
  $("sd-"+c.id).onclick = () => send(c);
  document.querySelectorAll("#card-"+c.id+" .q-btn").forEach(b => b.onclick = () => send(c, c.quick[b.dataset.k][1]));
  if(c.file){
    const z = $("up-"+c.id), fi = $("fi-"+c.id);
    z.onclick = () => fi.click();
    fi.onchange = () => { if(fi.files[0]) upload(c, fi.files[0]); };
    z.ondragover = e => { e.preventDefault(); z.classList.add("drag"); };
    z.ondragleave = () => z.classList.remove("drag");
    z.ondrop = e => { e.preventDefault(); z.classList.remove("drag"); if(e.dataTransfer.files[0]) upload(c, e.dataTransfer.files[0]); };
    $("fb-"+c.id).querySelector("button").onclick = () => { ST[c.id].file=""; fi.value=""; $("fb-"+c.id).classList.add("hidden"); };
  }
  if(c.sheet){
    $("sb-"+c.id).onclick = async () => {
      const url = $("su-"+c.id).value.trim(); if(!url) return;
      try{
        const d = await call("/api/connect-sheet", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({url})});
        if(d.status === "ok"){ ST[c.id].sheet = d.content; $("sh-"+c.id).classList.remove("hidden"); }
        else alert(d.error);
      }catch(e){ alert("Lỗi kết nối"); }
    };
    $("sh-"+c.id).querySelector("button").onclick = () => { ST[c.id].sheet=""; $("su-"+c.id).value=""; $("sh-"+c.id).classList.add("hidden"); };
  }
}
const grid = $("grid"), contact = $("contactCard");
CARDS.forEach(c => {
  ST[c.id] = {file:"", sheet:""};
  const el = document.createElement("div");
  el.className = "card "+c.cls; el.id = "card-"+c.id; el.innerHTML = cardHTML(c);
  grid.insertBefore(el, contact);
  wire(c); addMsg(c.id, "ai", c.hello);
});
</script>
</body>
</html>
'''

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
