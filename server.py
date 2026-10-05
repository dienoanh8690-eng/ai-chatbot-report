import os
import re
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from docx import Document
from openpyxl import Workbook, load_workbook
import requests
import io
import sys

# === SỬA LỖI MÃ HÓA TOÀN CỤC ===
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
CORS(app)

# === BIẾN MÔI TRƯỜNG ===
OPENROUTER_KEY = os.environ.get("OPENROUTER_KEY", "").strip()
print(f"🔑 OPENROUTER_KEY: {'✅ Đã có' if OPENROUTER_KEY else '❌ TRỐNG'}", flush=True)

# === CẤU HÌNH MÔ HÌNH ===
MODEL_CLAUDE = "anthropic/claude-3.5-sonnet"
MODEL_GEMINI = "google/gemini-2.0-flash-exp"
MODEL_DOLA  = "bytedance/dola-seed-2-0-pro"
MODEL_LLAMA = "meta-llama/llama-3.1-8b-instruct"

TIMEOUT = 60
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

MAC_DINH = "Trả lời bằng tiếng Việt rõ ràng, tự nhiên, dễ hiểu, chính xác. Sử dụng đầy đủ dấu thanh tiếng Việt theo chuẩn UTF-8."

# === GỌI AI — SỬA MÃ HÓA ===
def goi_ai_openrouter(ten_ai, model_id, prompt, he_thong=""):
    if not OPENROUTER_KEY:
        return ten_ai, None, "❌ Chưa đặt biến OPENROUTER_KEY trên Render"
    if not OPENROUTER_KEY.startswith("sk-or-v1-"):
        return ten_ai, None, "❌ Khóa OpenRouter sai định dạng"
    
    url = "https://openrouter.ai/api/v1/chat/completions"
    try:
        res = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {OPENROUTER_KEY}",
                "Content-Type": "application/json; charset=utf-8",
                "HTTP-Referer": "https://ai-chatbot-bao-cao.onrender.com/",
                "X-Title": "All Thủy Điện",
            },
            json={
                "model": model_id,
                "messages": [
                    {"role": "system", "content": he_thong or MAC_DINH},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.7,
                "max_tokens": 2048,
            },
            timeout=TIMEOUT,
        )
        
        print(f"📡 {ten_ai} — Mã: {res.status_code}", flush=True)
        
        if res.status_code == 200:
            data = res.json()
            text = data["choices"][0]["message"]["content"].strip()
            return ten_ai, text, None
        elif res.status_code == 401:
            return ten_ai, None, "❌ Khóa không hợp lệ/hết hạn"
        else:
            return ten_ai, None, f"Lỗi {res.status_code}: {res.text[:150]}"
            
    except UnicodeEncodeError as e:
        return ten_ai, None, f"❌ Lỗi mã hóa ký tự — vui lòng thử lại: {str(e)}"
    except Exception as e:
        return ten_ai, None, f"❌ Lỗi kết nối: {str(e)[:80]}"

# === PHÂN BỐ AI ===
def goi_soan_thao(prompt, he_thong=""):
    ten, kq, loi = goi_ai_openrouter("🟣 Claude", MODEL_CLAUDE, prompt, he_thong)
    if kq: return ten, kq, loi
    return goi_ai_openrouter("🟢 Llama", MODEL_LLAMA, prompt, he_thong)

def goi_dau_thau(prompt, he_thong=""):
    ten, kq, loi = goi_ai_openrouter("🔵 Gemini", MODEL_GEMINI, prompt, he_thong)
    if kq: return ten, kq, loi
    return goi_ai_openrouter("🟢 Meta/Llama", MODEL_LLAMA, prompt, he_thong)

def goi_thiet_bi(prompt, he_thong=""):
    ten, kq, loi = goi_ai_openrouter("🟠 Dola", MODEL_DOLA, prompt, he_thong)
    if kq: return ten, kq, loi
    return goi_ai_openrouter("🟢 Meta/Llama", MODEL_LLAMA, prompt, he_thong)

def goi_phan_tich(prompt, he_thong=""):
    return goi_ai_openrouter("🔵 Gemini", MODEL_GEMINI, prompt, he_thong)

CHATS = {
    "doc": ("Chuyên gia soạn thảo văn bản chuẩn Việt Nam. Viết trang trọng, đúng thể thức, đầy đủ dấu tiếng Việt.", goi_soan_thao),
    "tender": ("Chuyên gia đấu thầu theo Luật Việt Nam. Hướng dẫn chi tiết từng bước, dùng tiếng Việt chuẩn.", goi_dau_thau),
    "equip": ("Chuyên gia quản lý thiết bị thủy điện. Phân tích, đề xuất bảo trì, thay thế. Dùng tiếng Việt chuẩn.", goi_thiet_bi),
    "data": ("Chuyên gia phân tích dữ liệu & lập báo cáo. Tóm tắt, dùng bảng, ngắn gọn, chuẩn tiếng Việt.", goi_phan_tich),
}

# === XUẤT FILE ===
def _sach(dong):
    return re.sub(r"[*#`]+", "", dong).strip()

def tao_word(nd):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
        doc = Document()
        doc.add_heading("BÁO CÁO", 0)
        doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        for d in nd.split("\n"):
            s = _sach(d)
            if s and not re.fullmatch(r"[|\-:\s]+", d):
                doc.add_paragraph(s)
        doc.save(os.path.join(RESULT_FOLDER, ten))
        return ten
    except Exception as e:
        print(f"Lỗi tạo Word: {e}", flush=True)
        return ""

def tao_excel(nd):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
        wb = Workbook()
        ws = wb.active
        ws["A1"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        hang = 3
        for d in nd.split("\n"):
            if not d.strip(): continue
            if d.startswith("|") and re.fullmatch(r"[|\-:\s]+", d): continue
            if d.startswith("|"):
                for i, c in enumerate(d.strip("|").split("|"), 1):
                    ws.cell(row=hang, column=i, value=_sach(c))
            else:
                ws.cell(row=hang, column=1, value=_sach(d))
            hang += 1
        wb.save(os.path.join(RESULT_FOLDER, ten))
        return ten
    except Exception as e:
        print(f"Lỗi tạo Excel: {e}", flush=True)
        return ""

def doc_google_sheet(url):
    try:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
        if not m: return None, "Link không hợp lệ"
        r = requests.get(f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv", timeout=15)
        if r.status_code == 200:
            r.encoding = "utf-8"
            return r.text, None
        return None, "Vui lòng chia sẻ Sheet → Quyền: Bất kỳ ai có đường liên kết"
    except Exception as e: return None, f"Lỗi: {str(e)[:60]}"

# === API ===
@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Không có tệp"}), 400
    f = request.files["file"]
    ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
    if ext not in ("xlsx", "csv", "txt"):
        return jsonify({"error": "Chỉ hỗ trợ .xlsx, .csv, .txt"}), 400
    path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex[:10]}.{ext}")
    f.save(path)
    nd = ""
    try:
        if ext == "xlsx":
            wb = load_workbook(path, data_only=True, read_only=True)
            for ws in wb.worksheets:
                nd += f"=== {ws.title} ===\n"
                for row in ws.iter_rows(values_only=True):
                    nd += " | ".join("" if c is None else str(c) for c in row) + "\n"
                    if len(nd) > 6000: break
                if len(nd) > 6000: break
            wb.close()
        else:
            with open(path, "r", encoding="utf-8-sig", errors="replace") as fh:
                nd = fh.read()
    except Exception as e:
        nd = f"Lỗi đọc tệp: {e}"
    return jsonify({"status": "ok", "name": f.filename, "content": nd[:6000]})

@app.route("/api/connect-sheet", methods=["POST"])
def connect_sheet():
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()
    if not url: return jsonify({"error": "Dán link Google Sheets"}), 400
    nd, loi = doc_google_sheet(url)
    if loi: return jsonify({"error": loi}), 400
    return jsonify({"status": "ok", "content": nd[:6000]})

@app.route("/api/chat/<kind>", methods=["POST"])
def chat(kind):
    if kind not in CHATS:
        return jsonify({"reply": "Chức năng không tồn tại"}), 404
    data = request.get_json(silent=True) or {}
    msg = (data.get("message") or "").strip()
    ctx = "\n---\n".join(x for x in [(data.get("file_content") or "")[:6000],
                                      (data.get("sheet_content") or "")[:6000]] if x.strip())
    if not msg and not ctx:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải dữ liệu!"})
    
    prompt = f"{msg or 'Phân tích dữ liệu sau'}\n\nDữ liệu:\n{ctx}" if ctx else msg
    he_thong, goi_ham = CHATS[kind]
    
    ten, kq, loi = goi_ham(prompt, he_thong)
    
    out = {"reply": "", "word": "", "excel": ""}
    if kq:
        out["reply"] = f"✅ [{ten}]\n{kq}"
        if kind == "data":
            w, x = tao_word(kq), tao_excel(kq)
            out["word"] = f"/download/{w}" if w else ""
            out["excel"] = f"/download/{x}" if x else ""
    else:
        out["reply"] = f"❌ [{ten}]\n{loi}"
    return jsonify(out)

@app.route("/download/<ten>")
def download(ten):
    return send_from_directory(RESULT_FOLDER, os.path.basename(ten), as_attachment=True)

@app.route("/")
def index():
    return TRANG_CHU

# === GIAO DIỆN ĐÃ SỬA ===
TRANG_CHU = '''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Thủy Điện — Hệ thống hỗ trợ</title>
<style>
:root{--g100:#f1f5f9;--g200:#e2e8f0;--c-doc:#2563eb;--c-tender:#16a34a;--c-equip:#9333ea;--c-data:#f59e0b;--c-info:#ec4899}
*{margin:0;padding:0;box-sizing:border-box;font-family:system-ui,-apple-system,sans-serif}
body{background:linear-gradient(135deg,#f0f7ff,#faf5ff);min-height:100vh;padding:20px}
.header{text-align:center;margin-bottom:24px}
.header h1{font-size:24px;color:#1e293b;margin-bottom:4px}
.header p{color:#64748b;font-size:13px}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;max-width:2000px;margin:0 auto}
.row2{grid-column:1/-1;display:grid;grid-template-columns:2fr 1fr;gap:20px;max-width:2000px;margin:0 auto}
@media(max-width:1200px){.grid{grid-template-columns:repeat(2,1fr)}.row2{grid-template-columns:1fr}}
@media(max-width:768px){.grid,.row2{grid-template-columns:1fr}}
.card{background:#fff;border-radius:16px;padding:20px;box-shadow:0 4px 16px rgba(0,0,0,.05);display:flex;flex-direction:column;min-height:600px;border-top:4px solid var(--c)}
.c-doc{--c:var(--c-doc);--cl:#dbeafe}
.c-tender{--c:var(--c-tender);--cl:#dcfce7}
.c-equip{--c:var(--c-equip);--cl:#f3e8ff}
.c-data{--c:var(--c-data);--cl:#fef3c7}
.c-info{--c:var(--c-info);--cl:#fce7f3}
.card-head{display:flex;align-items:center;gap:8px;margin-bottom:12px;padding-bottom:10px;border-bottom:2px solid var(--cl)}
.card-icon{font-size:20px}.card-title{font-size:15px;font-weight:700;color:var(--c)}
.card-ai{font-size:10px;color:#94a3b8;margin-left:auto;background:var(--cl);padding:2px 8px;border-radius:10px;white-space:nowrap}
.upload-zone{border:2px dashed #cbd5e1;border-radius:12px;padding:12px;text-align:center;cursor:pointer;margin-bottom:10px;transition:.2s}
.upload-zone:hover,.upload-zone.drag{border-color:var(--c);background:var(--cl)}
.file-bar{display:flex;align-items:center;gap:8px;padding:8px 12px;border-radius:8px;margin-bottom:10px;font-size:12px;background:var(--cl)}
.file-bar button{margin-left:auto;background:none;border:none;font-size:16px;cursor:pointer;color:#dc2626}
.sheet-bar{display:flex;gap:8px;margin-bottom:10px}
.sheet-bar input{flex:1;padding:10px 12px;border:1px solid #e2e8f0;border-radius:8px;font-size:13px}
.sheet-bar button{padding:10px 14px;border:none;border-radius:8px;background:var(--c);color:#fff;cursor:pointer;font-weight:600;font-size:13px}
.quick-btns{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:12px}
.q-btn{padding:9px 10px;border:1px solid #e2e8f0;border-radius:8px;background:#fff;cursor:pointer;font-size:11.5px;text-align:left;transition:.2s}
.q-btn:hover{border-color:var(--c);background:var(--cl)}
.chat-area{flex:1;overflow-y:auto;padding:4px;margin-bottom:10px;min-height:220px;max-height:300px}
.msg{margin-bottom:12px;max-width:98%}.msg.user{margin-left:auto}
.bubble{padding:12px 14px;border-radius:16px;line-height:1.55;font-size:13.5px;white-space:pre-wrap}
.msg.user .bubble{background:linear-gradient(135deg,#dbeafe,#e0e7ff);border-bottom-right-radius:6px}
.msg.ai .bubble{background:#f8fafc;border-bottom-left-radius:6px}
.bubble.err{background:#fef2f2;border-left:3px solid #ef4444}
.dl-group{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap}
.dl-btn{padding:7px 14px;border-radius:16px;text-decoration:none;font-size:12px;font-weight:600;display:inline-block}
.dl-word{background:#dbeafe;color:#1d4ed8}.dl-excel{background:#dcfce7;color:#15803d}
.input-row{display:flex;gap:8px;align-items:flex-end;margin-top:10px}
textarea{flex:1;min-height:44px;max-height:100px;padding:11px 16px;border:1px solid #e2e8f0;border-radius:22px;font-size:14px;resize:none;outline:none;transition:.2s}
textarea:focus{border-color:var(--c);box-shadow:0 0 0 3px var(--cl)}
.send-btn{width:42px;height:42px;border-radius:50%;border:none;color:#fff;background:var(--c);cursor:pointer;font-size:17px;transition:.2s;flex-shrink:0}
.send-btn:hover{opacity:.9}
.send-btn:disabled{opacity:.5;cursor:not-allowed}
.hidden{display:none!important}
.contact-row{display:flex;align-items:flex-start;gap:8px;padding:7px 0;border-bottom:1px solid #f1f5f9;font-size:13.5px}
.contact-label{font-weight:600;color:var(--c);min-width:85px}
a{color:var(--c);text-decoration:none}
</style>
</head>
<body>
<div class="header">
<h1>⚡ All Thủy Điện — Hệ thống hỗ trợ</h1>
<p>Soạn thảo · Đấu thầu · Quản lý thiết bị · Phân tích dữ liệu</p>
</div>

<div class="grid">
<!-- CỘT 1: Soạn thảo văn bản -->
<div class="card c-doc">
<div class="card-head"><span class="card-icon">✍️</span><h3 class="card-title">Soạn thảo văn bản</h3><span class="card-ai">Claude → Llama</span></div>
<div class="quick-btns">
<button class="q-btn" data-preset="Soạn thảo công văn gửi cấp trên về tiến độ thực hiện dự án">Công văn hành chính</button>
<button class="q-btn" data-preset="Soạn thảo hợp đồng cung cấp dịch vụ kỹ thuật theo chuẩn nhà nước">Hợp đồng & Thỏa thuận</button>
<button class="q-btn" data-preset="Viết báo cáo công việc tháng, nêu kết quả, khó khăn, kế hoạch tháng sau">Báo cáo công việc</button>
<button class="q-btn" data-preset="Soạn thảo biên bản cuộc họp đánh giá tiến độ công trình">Biên bản cuộc họp</button>
</div>
<div class="chat-area" id="ch-doc">
<div class="msg ai"><div class="bubble">👋 Xin chào! Tôi dùng <strong>Claude</strong> chuyên soạn thảo văn bản chuẩn Việt Nam. Bạn cần viết gì?</div></div>
</div>
<div class="input-row"><textarea id="in-doc" placeholder="Nội dung cần soạn thảo..."></textarea><button class="send-btn" id="sd-doc" title="Gửi">➤</button></div>
</div>

<!-- CỘT 2: Quy trình đấu thầu -->
<div class="card c-tender">
<div class="card-head"><span class="card-icon">🏆</span><h3 class="card-title">Quy trình đấu thầu</h3><span class="card-ai">Gemini → Meta</span></div>
<div class="quick-btns">
<button class="q-btn" data-preset="Trình bày toàn bộ quy trình đấu thầu từ chuẩn bị đến nghiệm thu theo Luật Đấu thầu">Toàn bộ quy trình</button>
<button class="q-btn" data-preset="Liệt kê chi tiết hồ sơ mời thầu cần chuẩn bị cho gói cung cấp thiết bị">Hồ sơ mời thầu</button>
<button class="q-btn" data-preset="Giải thích các điều khoản pháp lý quan trọng cần lưu ý khi đấu thầu">Pháp lý & Lưu ý</button>
<button class="q-btn" data-preset="Danh sách các biểu mẫu thông dụng trong quá trình đấu thầu">Biểu mẫu thông dụng</button>
</div>
<div class="chat-area" id="ch-tender">
<div class="msg ai"><div class="bubble">👋 Tôi dùng <strong>Gemini + Meta</strong> hướng dẫn theo Luật Đấu thầu Việt Nam. Cần hỗ trợ bước nào?</div></div>
</div>
<div class="input-row"><textarea id="in-tender" placeholder="Hỏi về quy trình đấu thầu..."></textarea><button class="send-btn" id="sd-tender" title="Gửi">➤</button></div>
</div>

<!-- CỘT 3: Quản lý thiết bị -->
<div class="card c-equip">
<div class="card-head"><span class="card-icon">🔧</span><h3 class="card-title">Quản lý thiết bị</h3><span class="card-ai">Dola → Meta</span></div>
<div class="upload-zone" id="up-equip"><p>📎 Tải danh sách thiết bị (.xlsx, .csv, .txt)</p></div>
<input type="file" id="fi-equip" accept=".xlsx,.csv,.txt" class="hidden">
<div class="file-bar hidden" id="fb-equip"><span></span><button>✕</button></div>
<div class="quick-btns">
<button class="q-btn" data-preset="Phân loại thiết bị nhà máy thủy điện theo chức năng và hệ thống">Phân loại thiết bị</button>
<button class="q-btn" data-preset="Lập kế hoạch bảo trì định kỳ cho thiết bị chính của nhà máy">Kế hoạch bảo trì</button>
<button class="q-btn" data-preset="Đánh giá tình trạng thiết bị, xác định rủi ro và đề xuất xử lý">Đánh giá tình trạng</button>
<button class="q-btn" data-preset="Phân tích tuổi thọ thiết bị, đề xuất kế hoạch thay thế hợp lý">Tuổi thọ & Thay thế</button>
</div>
<div class="chat-area" id="ch-equip">
<div class="msg ai"><div class="bubble">👋 Tôi dùng <strong>Dola + Meta</strong> phân tích thiết bị. Tải danh sách hoặc nhập yêu cầu nhé!</div></div>
</div>
<div class="input-row"><textarea id="in-equip" placeholder="Nhập yêu cầu quản lý thiết bị..."></textarea><button class="send-btn" id="sd-equip" title="Gửi">➤</button></div>
</div>
</div>

<div class="row2">
<!-- DỮ LIỆU & BÁO CÁO -->
<div class="card c-data">
<div class="card-head"><span class="card-icon">📊</span><h3 class="card-title">Xử lý dữ liệu & Tạo báo cáo</h3><span class="card-ai">Gemini</span></div>
<div class="upload-zone" id="up-data"><p>📎 Tải tệp dữ liệu (.xlsx, .csv, .txt)</p></div>
<input type="file" id="fi-data" accept=".xlsx,.csv,.txt" class="hidden">
<div class="file-bar hidden" id="fb-data"><span></span><button>✕</button></div>
<div class="sheet-bar"><input type="text" id="su-data" placeholder="🔗 Dán link Google Sheets..."><button id="sb-data">Kết nối</button></div>
<div class="file-bar hidden" id="sh-data"><span>✅ Đã kết nối Sheets</span><button>✕</button></div>
<div class="quick-btns">
<button class="q-btn" data-preset="Tóm tắt số liệu chính, nêu nhận xét xu hướng từ dữ liệu">Tóm tắt dữ liệu</button>
<button class="q-btn" data-preset="Tính tổng, trung bình, tỷ lệ, so sánh các chỉ số">Tính tổng hợp</button>
<button class="q-btn" data-preset="Viết báo cáo đầy đủ có cấu trúc, kết luận và đề xuất">Báo cáo đầy đủ</button>
<button class="q-btn" data-preset="Đánh giá dữ liệu, nêu nhận xét chính và đề xuất cải tiến">Nhận xét & Đề xuất</button>
</div>
<div class="chat-area" id="ch-data">
<div class="msg ai"><div class="bubble">👋 Tôi dùng <strong>Gemini</strong> phân tích dữ liệu và xuất Word/Excel. Tải tệp hoặc dán link Sheets nhé!</div></div>
</div>
<div class="input-row"><textarea id="in-data" placeholder="Nhập yêu cầu phân tích..."></textarea><button class="send-btn" id="sd-data" title="Gửi">➤</button></div>
</div>

<!-- THÔNG TIN LIÊN HỆ -->
<div class="card c-info">
<div class="card-head"><span class="card-icon">📌</span><h3 class="card-title">Thông tin liên hệ</h3></div>
<div style="flex:1">
<div class="contact-row"><span>👤</span><span class="contact-label">Admin:</span><span>Chotvjp</span></div>
<div class="contact-row"><span>📞</span><span class="contact-label">ĐT:</span><a href="tel:0973020486">0973020486</a></div>
<div class="contact-row"><span>✉️</span><span class="contact-label">Email:</span><a href="mailto:hoangdien86ncc@gmail.com">hoangdien86ncc@gmail.com</a></div>
<div class="contact-row"><span>🏢</span><span class="contact-label">Đơn vị:</span><span>Công ty Cổ phần Thủy điện Nậm Chiến</span></div>
<div class="contact-row"><span>📍</span><span class="contact-label">Địa chỉ:</span><span>TK5 - Mường La - Sơn La</span></div>
<div class="contact-row"><span>🌐</span><span class="contact-label">Web:</span><a href="https://namchien.vn" target="_blank">namchien.vn</a></div>
</div>
<div style="margin-top:auto;padding-top:14px;text-align:center;font-size:11.5px;color:#9ca3af">
<p>© 2026 — Hệ thống hỗ trợ công việc nội bộ</p>
</div>
</div>
</div>

<script>
const $ = id => document.getElementById(id);
const ST = {equip:{file:""}, data:{file:"", sheet:""}};

function addMsg(kind, type, text, d){
  const m = document.createElement("div"); m.className = "msg "+type;
  const b = document.createElement("div"); b.className = "bubble"+(text.includes("❌")?" err":"");
  b.innerHTML = text.replace(/\n/g,"<br>"); m.appendChild(b);
  if(d && (d.word || d.excel)){
    const g = document.createElement("div"); g.className = "dl-group";
    if(d.word) g.innerHTML += '<a class="dl-btn dl-word" href="'+d.word+'" download>📄 Tải Word</a>';
    if(d.excel) g.innerHTML += '<a class="dl-btn dl-excel" href="'+d.excel+'" download>📊 Tải Excel</a>';
    b.appendChild(g);
  }
  const c = $("ch-"+kind); c.appendChild(m); c.scrollTop = c.scrollHeight;
}

async function send(kind, preset){
  const i = $("in-"+kind), s = ST[kind] || {};
  const m = (preset !== undefined ? preset : i.value).trim();
  if(!m && !s.file && !s.sheet) return;
  
  addMsg(kind, "user", m || "Phân tích dữ liệu");
  i.value = "";
  
  const btn = $("sd-"+kind); 
  btn.disabled = true; btn.textContent = "⏳";
  
  try{
    const r = await fetch("/api/chat/"+kind, {
      method:"POST",
      headers:{"Content-Type":"application/json; charset=utf-8"},
      body: JSON.stringify({
        message: m,
        file_content: (s.file || "").substring(0, 6000),
        sheet_content: (s.sheet || "").substring(0, 6000)
      })
    });
    const d = await r.json();
    addMsg(kind, "ai", d.reply || "❌ Không có phản hồi", d);
  }catch(e){
    addMsg(kind, "ai", "❌ Lỗi kết nối — vui lòng chờ 30-60 giây thử lại nhé!");
  }finally{
    btn.disabled = false; btn.textContent = "➤";
  }
}

async function upload(kind, f){
  const fd = new FormData(); fd.append("file", f);
  try{
    const r = await fetch("/api/upload", {method:"POST", body:fd});
    const d = await r.json();
    if(d.status === "ok"){
      ST[kind].file = d.content;
      $("fb-"+kind).querySelector("span").textContent = "📎 "+d.name;
      $("fb-"+kind).classList.remove("hidden");
    } else alert(d.error);
  }catch(e){ alert("Không tải được tệp"); }
}

// === SOẠN THẢO ===
$("in-doc").addEventListener("keydown", e => { if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); send("doc"); } });
$("sd-doc").onclick = () => send("doc");
document.querySelectorAll(".card.c-doc .q-btn").forEach(b => b.onclick = () => send("doc", b.dataset.preset));

// === ĐẤU THẦU ===
$("in-tender").addEventListener("keydown", e => { if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); send("tender"); } });
$("sd-tender").onclick = () => send("tender");
document.querySelectorAll(".card.c-tender .q-btn").forEach(b => b.onclick = () => send("tender", b.dataset.preset));

// === THIẾT BỊ ===
$("in-equip").addEventListener("keydown", e => { if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); send("equip"); } });
$("sd-equip").onclick = () => send("equip");
$("up-equip").onclick = () => $("fi-equip").click();
$("fi-equip").onchange = () => { if($("fi-equip").files[0]) upload("equip", $("fi-equip").files[0]); };
$("up-equip").ondragover = e => { e.preventDefault(); $("up-equip").classList.add("drag"); };
$("up-equip").ondragleave = () => $("up-equip").classList.remove("drag");
$("up-equip").ondrop = e => { e.preventDefault(); $("up-equip").classList.remove("drag"); if(e.dataTransfer.files[0]) upload("equip", e.dataTransfer.files[0]); };
$("fb-equip").querySelector("button").onclick = () => { ST.equip.file=""; $("fi-equip").value=""; $("fb-equip").classList.add("hidden"); };
document.querySelectorAll(".card.c-equip .q-btn").forEach(b => b.onclick = () => send("equip", b.dataset.preset));

// === DỮ LIỆU ===
$("in-data").addEventListener("keydown", e => { if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); send("data"); } });
$("sd-data").onclick = () => send("data");
$("up-data").onclick = () => $("fi-data").click();
$("fi-data").onchange = () => { if($("fi-data").files[0]) upload("data", $("fi-data").files[0]); };
$("up-data").ondragover = e => { e.preventDefault(); $("up-data").classList.add("drag"); };
$("up-data").ondragleave = () => $("up-data").classList.remove("drag");
$("up-data").ondrop = e => { e.preventDefault(); $("up-data").classList.remove("drag"); if(e.dataTransfer.files[0]) upload("data", e.dataTransfer.files[0]); };
$("fb-data").querySelector("button").onclick = () => { ST.data.file=""; $("fi-data").value=""; $("fb-data").classList.add("hidden"); };
$("sb-data").onclick = async () => {
  const url = $("su-data").value.trim(); if(!url) return;
  try{
    const r = await fetch("/api/connect-sheet", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({url})});
    const d = await r.json();
    if(d.status === "ok"){ ST.data.sheet = d.content; $("sh-data").classList.remove("hidden"); }
    else alert(d.error);
  }catch(e){ alert("Lỗi kết nối"); }
};
$("sh-data").querySelector("button").onclick = () => { ST.data.sheet=""; $("su-data").value=""; $("sh-data").classList.add("hidden"); };
document.querySelectorAll(".card.c-data .q-btn").forEach(b => b.onclick = () => send("data", b.dataset.preset));
</script>
</body>
</html>
'''

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    print(f"🚀 Server chạy trên cổng {port}", flush=True)
    app.run(host="0.0.0.0", port=port)
