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
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # tối đa 10MB/tệp
CORS(app)

# ==================================================
# CẤU HÌNH — đặt trong Render → Environment
# ==================================================
def env(name, default=""):
    return os.environ.get(name, default).strip()

AI_API_KEY = env("AI_API_KEY")                      # AIML API
AI_MODEL = env("AI_MODEL", "bytedance/dola-seed-2-0-pro")
GEMINI_API_KEY = env("GEMINI_API_KEY")
GEMINI_MODEL = env("GEMINI_MODEL", "gemini-2.5-flash")
GROQ_API_KEY = env("GROQ_API_KEY")
GROQ_MODEL = env("GROQ_MODEL", "llama-3.1-8b-instant")
OPENAI_API_KEY = env("OPENAI_API_KEY")
OPENAI_MODEL = env("OPENAI_MODEL", "gpt-4o-mini")
CLAUDE_API_KEY = env("CLAUDE") or env("ANTHROPIC_API_KEY")
CLAUDE_MODEL = env("CLAUDE_MODEL", "claude-haiku-4-5-20251001")
APP_PASSWORD = env("APP_PASSWORD")                  # tuỳ chọn: mã truy cập chung

TIMEOUT = 20
UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)
MAC_DINH = "Trả lời bằng tiếng Việt rõ ràng, tự nhiên."


# ==================================================
# GỌI TỪNG AI — mỗi hàm trả về (tên, nội dung, lỗi)
# ==================================================
def _openai_style(ten, url, key, model, prompt, he_thong):
    if not key:
        return ten, None, "chưa đặt khóa API"
    try:
        res = requests.post(
            url,
            headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
            json={
                "model": model,
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
            text = res.json()["choices"][0]["message"]["content"]
            if text and text.strip():
                return ten, text, None
            return ten, None, "phản hồi rỗng"
        print(f"[{ten}] HTTP {res.status_code}: {res.text[:300]}", flush=True)
        return ten, None, f"HTTP {res.status_code}"
    except Exception as e:
        print(f"[{ten}] Exception: {e}", flush=True)
        return ten, None, f"lỗi kết nối ({type(e).__name__})"


def goi_aiml(prompt, he_thong=""):
    return _openai_style("DOLA/AIML", "https://api.aimlapi.com/v1/chat/completions",
                         AI_API_KEY, AI_MODEL, prompt, he_thong)


def goi_groq(prompt, he_thong=""):
    return _openai_style("Llama/Groq", "https://api.groq.com/openai/v1/chat/completions",
                         GROQ_API_KEY, GROQ_MODEL, prompt, he_thong)


def goi_gpt(prompt, he_thong=""):
    return _openai_style("GPT", "https://api.openai.com/v1/chat/completions",
                         OPENAI_API_KEY, OPENAI_MODEL, prompt, he_thong)


def goi_gemini(prompt, he_thong=""):
    ten = "Gemini"
    if not GEMINI_API_KEY:
        return ten, None, "chưa đặt khóa API"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    try:
        res = requests.post(
            url,
            headers={"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"},
            json={
                "systemInstruction": {"parts": [{"text": he_thong or MAC_DINH}]},
                "contents": [{"role": "user", "parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.7, "maxOutputTokens": 2048},
            },
            timeout=TIMEOUT,
        )
        if res.status_code == 200:
            cands = res.json().get("candidates") or []
            if cands:
                parts = cands[0].get("content", {}).get("parts", [])
                text = "".join(p.get("text", "") for p in parts).strip()
                if text:
                    return ten, text, None
            return ten, None, "phản hồi rỗng/bị chặn"
        print(f"[{ten}] HTTP {res.status_code}: {res.text[:300]}", flush=True)
        return ten, None, f"HTTP {res.status_code}"
    except Exception as e:
        print(f"[{ten}] Exception: {e}", flush=True)
        return ten, None, f"lỗi kết nối ({type(e).__name__})"


def goi_claude(prompt, he_thong=""):
    ten = "Claude"
    if not CLAUDE_API_KEY:
        return ten, None, "chưa đặt khóa API"
    try:
        res = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": CLAUDE_API_KEY,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json",
            },
            json={
                "model": CLAUDE_MODEL,
                "max_tokens": 2048,
                "system": he_thong or MAC_DINH,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=TIMEOUT,
        )
        if res.status_code == 200:
            blocks = res.json().get("content", [])
            text = "".join(b.get("text", "") for b in blocks if b.get("type") == "text").strip()
            if text:
                return ten, text, None
            return ten, None, "phản hồi rỗng"
        print(f"[{ten}] HTTP {res.status_code}: {res.text[:300]}", flush=True)
        return ten, None, f"HTTP {res.status_code}"
    except Exception as e:
        print(f"[{ten}] Exception: {e}", flush=True)
        return ten, None, f"lỗi kết nối ({type(e).__name__})"


def goi_theo_danh_sach(prompt, danh_sach_ham, he_thong=""):
    """Thử lần lượt từng AI. Trả về dict: ok, reply (hiển thị), plain (nội dung thuần)."""
    loi_tong = []
    for ham in danh_sach_ham:
        ten, kq, loi = ham(prompt, he_thong)
        if kq:
            return {"ok": True, "reply": f"✅ [{ten}]\n{kq}", "plain": kq}
        loi_tong.append(f"{ten}: {loi}")
    reply = "❌ Tất cả AI đều không trả lời:\n" + "\n".join(f"× {x}" for x in loi_tong)
    return {"ok": False, "reply": reply, "plain": ""}


# Mỗi chức năng: (lời nhắc hệ thống, thứ tự AI thử)
CHATS = {
    "general": ("Bạn là trợ lý AI tổng hợp, thông minh, hữu ích. Trả lời rõ ràng, dễ hiểu, bằng tiếng Việt.",
                [goi_gemini, goi_aiml, goi_groq, goi_gpt, goi_claude]),
    "data": ("Bạn là chuyên gia phân tích dữ liệu và lập báo cáo. Trả lời bằng tiếng Việt, tóm tắt số liệu quan trọng, "
             "dùng bảng khi phù hợp.",
             [goi_gemini, goi_aiml, goi_groq, goi_gpt]),
    "doc": ("Bạn là chuyên gia soạn thảo văn bản hành chính, hợp đồng, thư từ. Viết chuẩn mực, đúng thể thức Việt Nam.",
            [goi_gpt, goi_claude, goi_gemini, goi_aiml]),
    "tender": ("Bạn là chuyên gia tư vấn quy trình đấu thầu theo pháp luật Việt Nam. Hướng dẫn chi tiết từng bước, "
               "hồ sơ, lưu ý pháp lý; nhắc người dùng đối chiếu văn bản pháp luật hiện hành.",
               [goi_gemini, goi_groq, goi_aiml, goi_claude]),
    "equip": ("Bạn là chuyên gia quản lý thiết bị nhà máy thủy điện. Phân loại, theo dõi tình trạng, đề xuất bảo trì, "
              "tính tuổi thọ. Trả lời bằng tiếng Việt.",
              [goi_gemini, goi_groq, goi_aiml]),
}


# ==================================================
# TẠO TỆP WORD & EXCEL
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
            if not d:
                continue
            if d.startswith("|"):  # dòng bảng markdown → tách thành ô
                if re.fullmatch(r"[|\-:\s]+", d):
                    continue
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
        return None, "❌ Không đọc được Sheet → Chia sẻ: «Bất kỳ ai có đường liên kết» (Người xem)"
    except Exception as e:
        return None, f"❌ Lỗi đọc Sheet: {e}"


# ==================================================
# BẢO VỆ BẰNG MÃ TRUY CẬP (nếu đặt APP_PASSWORD)
# ==================================================
@app.before_request
def kiem_tra_ma():
    if APP_PASSWORD and request.path.startswith("/api/"):
        if request.headers.get("X-Access-Code", "") != APP_PASSWORD:
            return jsonify({"error": "Cần mã truy cập", "reply": "❌ Sai hoặc thiếu mã truy cập."}), 401


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
                    if len(noi_dung) > 6000:
                        break
                if len(noi_dung) > 6000:
                    break
            wb.close()
        else:
            with open(duong_dan, "r", encoding="utf-8-sig", errors="ignore") as fh:
                noi_dung = fh.read()
    except Exception as e:
        noi_dung = f"(Không đọc được nội dung: {e})"
    return jsonify({"status": "ok", "name": ten_goc, "content": noi_dung[:6000]})


@app.route("/api/connect-sheet", methods=["POST"])
def connect_sheet():
    url = ((request.get_json(silent=True) or {}).get("url") or "").strip()
    if not url:
        return jsonify({"error": "Vui lòng dán link Google Sheets"}), 400
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
    prompt = f"{msg or 'Hãy phân tích dữ liệu sau.'}\n\nDữ liệu đính kèm:\n{ctx}" if ctx else msg
    he_thong, danh_sach = CHATS[kind]
    kq = goi_theo_danh_sach(prompt, danh_sach, he_thong)
    out = {"reply": kq["reply"], "word": "", "excel": ""}
    if kind == "data" and kq["ok"]:
        w, x = tao_word(kq["plain"]), tao_excel(kq["plain"])
        out["word"] = f"/download/{w}" if w else ""
        out["excel"] = f"/download/{x}" if x else ""
    return jsonify(out)


@app.route("/download/<path:ten_file>")
def download(ten_file):
    ten_file = os.path.basename(ten_file)
    return send_from_directory(RESULT_FOLDER, ten_file, as_attachment=True)


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
<title>All Thủy Điện — Hệ thống hỗ trợ toàn diện</title>
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
.card{background:#fff;border-radius:20px;padding:24px;box-shadow:0 4px 12px rgba(0,0,0,.06);display:flex;flex-direction:column;height:720px;min-height:650px;border-top:4px solid var(--c)}
.c-gen{--c:#6366f1;--cl:#e0e7ff;--cb:#eef2ff}.c-p1{--c:#2563eb;--cl:#dbeafe;--cb:#eff6ff}
.c-p2{--c:#16a34a;--cl:#dcfce7;--cb:#f0fdf4}.c-p3{--c:#9333ea;--cl:#f3e8ff;--cb:#faf5ff}
.c-p4{--c:#f59e0b;--cl:#fef3c7;--cb:#fffbeb}.c-p5{--c:#ec4899;--cl:#fce7f3;--cb:#fdf2f8}
.card-head{display:flex;align-items:center;gap:10px;margin-bottom:18px;padding-bottom:14px;border-bottom:2px solid var(--cl);flex-wrap:wrap}
.card-icon{font-size:24px}.card-title{font-size:17px;font-weight:700;color:var(--c)}
.card-ai{font-size:11px;color:#94a3b8;margin-left:auto;background:var(--g100);padding:3px 8px;border-radius:12px}
.upload-zone{border:2px dashed var(--g200);border-radius:12px;padding:14px;text-align:center;cursor:pointer;margin-bottom:12px}
.upload-zone:hover,.upload-zone.drag{border-color:var(--c);background:var(--cb)}
.upload-zone p{font-size:13px;color:var(--g600)}
.file-bar{display:flex;align-items:center;gap:8px;padding:10px 14px;border-radius:8px;margin-bottom:12px;font-size:13px;font-weight:500;background:var(--cl);color:var(--g800)}
.file-bar button{margin-left:auto;background:none;border:none;font-size:18px;cursor:pointer;color:#dc2626}
.sheet-bar{display:flex;gap:8px;margin-bottom:12px}
.sheet-bar input{flex:1;padding:10px 14px;border:1px solid var(--g200);border-radius:8px;font-size:13px;outline:none}
.sheet-bar button{padding:10px 16px;border:none;border-radius:8px;background:var(--c);color:#fff;cursor:pointer;font-weight:600;font-size:13px}
.quick-btns{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px}
.q-btn{padding:10px 12px;border:1px solid var(--g200);border-radius:8px;background:#fff;cursor:pointer;font-size:12px;text-align:left;line-height:1.4}
.q-btn:hover{border-color:var(--c);background:var(--cb)}
.chat-area{flex:1;overflow-y:auto;padding:4px;margin-bottom:14px;min-height:200px}
.msg{margin-bottom:16px;max-width:98%}.msg.user{margin-left:auto}
.bubble{padding:14px 18px;border-radius:18px;line-height:1.6;font-size:14px;white-space:pre-wrap;word-break:break-word}
.msg.user .bubble{background:linear-gradient(135deg,#dbeafe,#e0e7ff);border-bottom-right-radius:8px;color:#1e3a8a}
.msg.ai .bubble{background:var(--g100);border-bottom-left-radius:8px;color:var(--g800)}
.bubble.warn{background:#fef3c7;border-left:3px solid #f59e0b;color:#92400e}
.bubble.err{background:#fee2e2;border-left:3px solid #ef4444;color:#b91c1c}
.dl-group{display:flex;gap:10px;margin-top:12px;flex-wrap:wrap}
.dl-btn{padding:8px 16px;border-radius:20px;text-decoration:none;font-size:13px;font-weight:600}
.dl-word{background:#dbeafe;color:#1d4ed8}.dl-excel{background:#dcfce7;color:#15803d}
.input-row{display:flex;gap:10px;align-items:flex-end}
textarea{flex:1;min-height:48px;max-height:120px;padding:12px 18px;border:1px solid var(--g200);border-radius:24px;font-size:14px;resize:none;outline:none;line-height:1.5}
textarea:focus{border-color:var(--c)}
.send-btn{width:44px;height:44px;border-radius:50%;border:none;color:#fff;background:var(--c);cursor:pointer;font-size:18px;flex-shrink:0}
.send-btn:disabled{opacity:.5;cursor:not-allowed}
.hidden{display:none!important}
.contact-row{display:flex;align-items:flex-start;gap:10px;padding:8px 0;border-bottom:1px solid var(--g100);font-size:14px}
.contact-row:last-child{border-bottom:none}
.contact-label{font-weight:600;color:var(--c);min-width:90px;flex-shrink:0}
.contact-row a{color:var(--c);text-decoration:none}
</style>
</head>
<body>
<div class="header">
<h1>⚡ All Thủy Điện — Hệ thống hỗ trợ toàn diện</h1>
<p>Trò chuyện chung · Phân tích dữ liệu · Soạn thảo văn bản · Tư vấn đấu thầu · Quản lý thiết bị · Thông tin liên hệ</p>
</div>
<div class="grid" id="grid">
<div class="card c-p5" id="contactCard">
<div class="card-head"><span class="card-icon">📌</span><h3 class="card-title">Thông tin liên hệ</h3></div>
<div>
<div class="contact-row"><span>👤</span><span class="contact-label">Admin:</span><span>Chotvjp</span></div>
<div class="contact-row"><span>📞</span><span class="contact-label">Điện thoại:</span><a href="tel:0973020486">0973020486</a></div>
<div class="contact-row"><span>✉️</span><span class="contact-label">Email:</span><a href="mailto:hoangdien86ncc@gmail.com">hoangdien86ncc@gmail.com</a></div>
<div class="contact-row"><span>🏢</span><span class="contact-label">Bộ phận:</span><span>Quản lý kỹ thuật</span></div>
<div class="contact-row"><span>🏭</span><span class="contact-label">Đơn vị:</span><span>Công ty Cổ phần Thủy điện Nậm Chiến</span></div>
<div class="contact-row"><span>📍</span><span class="contact-label">Địa chỉ:</span><span>TK5 - Mường La - Sơn La</span></div>
<div class="contact-row"><span>🌐</span><span class="contact-label">Website:</span><a href="https://namchien.vn" target="_blank" rel="noopener">namchien.vn</a></div>
</div>
<div style="margin-top:auto;padding-top:16px;text-align:center;font-size:12px;color:#9ca3af">
<p>Hệ thống hỗ trợ công việc nội bộ</p><p>© 2026 Công ty Cổ phần Thủy điện Nậm Chiến</p>
</div>
</div>
</div>

<script>
const CARDS = [
 {id:"general",cls:"c-gen",icon:"💬",title:"Trò chuyện chung",ai:"Gemini → DOLA → Llama → GPT → Claude",
  hello:"👋 Xin chào! Tôi là trợ lý AI tổng hợp. Bạn có thể hỏi bất kỳ điều gì nhé!",ph:"Hỏi bất kỳ điều gì...",
  quick:[["🔌 Thủy điện cơ bản","Giải thích khái niệm về nhà máy thủy điện"],["📰 Tin tức & Công nghệ","Tóm tắt xu hướng công nghệ mới trong ngành điện"],["💡 Ý tưởng & Đề xuất","Đề xuất ý tưởng tối ưu hóa hiệu suất làm việc"],["⚖️ Pháp luật chung","Giải đáp thắc mắc chung về pháp luật"]]},
 {id:"data",cls:"c-p1",icon:"📊",title:"Xử lý dữ liệu & Tạo báo cáo",ai:"Gemini → DOLA → Llama → GPT",file:true,sheet:true,
  hello:"👋 Tải tệp, dán link Google Sheets hoặc nhập yêu cầu để bắt đầu phân tích nhé!",ph:"Nhập yêu cầu phân tích...",
  quick:[["📋 Sắp xếp dữ liệu","Sắp xếp và tóm tắt dữ liệu"],["💰 Tính tổng hợp","Tính tổng và phân tích số liệu"],["📑 Lập báo cáo","Lập báo cáo đầy đủ có cấu trúc"],["📈 Nhận xét & Đề xuất","Đánh giá xu hướng và đề xuất"]]},
 {id:"doc",cls:"c-p2",icon:"✍️",title:"Soạn thảo văn bản",ai:"GPT → Claude → Gemini → DOLA",
  hello:"👋 Tôi sẽ giúp bạn soạn thảo văn bản chuẩn mực, đúng thể thức Việt Nam. Bạn cần viết gì?",ph:"Bạn cần soạn thảo gì...?",
  quick:[["📝 Công văn","Soạn thảo công văn gửi cấp trên"],["📄 Hợp đồng","Soạn thảo hợp đồng mua bán thiết bị"],["📈 Báo cáo tiến độ","Viết báo cáo tiến độ thực hiện dự án"],["📋 Thư & Biên bản","Soạn thảo thư mời họp và biên bản"]]},
 {id:"tender",cls:"c-p3",icon:"🏆",title:"Quy trình đấu thầu",ai:"Gemini → Llama → DOLA → Claude",
  hello:"👋 Tôi hướng dẫn chi tiết theo quy định Việt Nam. Bạn cần hỗ trợ về bước nào?",ph:"Hỏi về quy trình đấu thầu...?",
  quick:[["📋 Toàn bộ quy trình","Giải thích toàn bộ quy trình đấu thầu"],["📑 Hồ sơ mời thầu","Danh mục hồ sơ cần chuẩn bị"],["⚖️ Pháp lý & Rủi ro","Lưu ý pháp lý và rủi ro thường gặp"],["📄 Biểu mẫu","Mẫu biểu mẫu thông dụng"]]},
 {id:"equip",cls:"c-p4",icon:"🔧",title:"Quản lý thiết bị",ai:"Gemini → Llama → DOLA",file:true,
  hello:"👋 Tải danh sách thiết bị hoặc nhập yêu cầu — tôi sẽ phân tích chi tiết nhé!",ph:"Nhập yêu cầu quản lý thiết bị...?",
  quick:[["📊 Phân loại thiết bị","Phân loại thiết bị theo nhóm"],["🛠️ Kế hoạch bảo trì","Đề xuất kế hoạch bảo trì định kỳ"],["⚠️ Đánh giá rủi ro","Đánh giá tình trạng và rủi ro"],["🔄 Tuổi thọ & Thay thế","Tính tuổi thọ và đề xuất thay thế"]]}
];

let CODE = sessionStorage.getItem("code") || "";
const $ = id => document.getElementById(id);
const ST = {};

function cardHTML(c){
  return '<div class="card-head"><span class="card-icon">'+c.icon+'</span><h3 class="card-title">'+c.title+'</h3><span class="card-ai">'+c.ai+'</span></div>'
  +(c.file?'<div class="upload-zone" id="up-'+c.id+'"><p>📎 Nhấn chọn hoặc kéo thả tệp (.xlsx, .csv, .txt)</p></div><input type="file" id="fi-'+c.id+'" accept=".xlsx,.csv,.txt" class="hidden"><div class="file-bar hidden" id="fb-'+c.id+'"><span></span><button>✕</button></div>':"")
  +(c.sheet?'<div class="sheet-bar"><input type="text" id="su-'+c.id+'" placeholder="🔗 Dán link Google Sheets..."><button id="sb-'+c.id+'">Kết nối</button></div><div class="file-bar hidden" id="sh-'+c.id+'"><span>✅ Google Sheets đã kết nối</span><button>✕</button></div>':"")
  +'<div class="quick-btns">'+c.quick.map((q,k)=>'<button class="q-btn" data-k="'+k+'">'+q[0]+'</button>').join("")+'</div>'
  +'<div class="chat-area" id="ch-'+c.id+'"></div>'
  +'<div class="input-row"><textarea id="in-'+c.id+'" placeholder="'+c.ph+'"></textarea><button class="send-btn" id="sd-'+c.id+'">➤</button></div>';
}

function addMsg(id, type, text, d){
  const m = document.createElement("div"); m.className = "msg "+type;
  const b = document.createElement("div"); b.className = "bubble"+(text.includes("❌")?" err":text.includes("⚠️")?" warn":"");
  b.textContent = text; m.appendChild(b);
  if(d && (d.word || d.excel)){
    const g = document.createElement("div"); g.className = "dl-group";
    if(d.word) g.innerHTML += '<a class="dl-btn dl-word" href="'+d.word+'" target="_blank">📄 Tải Word</a>';
    if(d.excel) g.innerHTML += '<a class="dl-btn dl-excel" href="'+d.excel+'" target="_blank">📊 Tải Excel</a>';
    b.appendChild(g);
  }
  const c = $("ch-"+id); c.appendChild(m); c.scrollTop = c.scrollHeight;
}

async function call(path, opts){
  opts.headers = Object.assign({"X-Access-Code": CODE}, opts.headers || {});
  const r = await fetch(path, opts);
  if(r.status === 401){
    const nhap = prompt("Nhập mã truy cập:");
    if(!nhap) throw new Error("Cần mã truy cập");
    CODE = nhap; sessionStorage.setItem("code", CODE);
    return call(path, opts);
  }
  return r.json();
}

async function send(c, preset){
  const i = $("in-"+c.id), s = ST[c.id];
  const m = (preset !== undefined ? preset : i.value).trim();
  if(!m && !s.file && !s.sheet) return;
  addMsg(c.id, "user", m || "Phân tích dữ liệu"); i.value = "";
  const btn = $("sd-"+c.id); btn.disabled = true; btn.textContent = "⏳";
  try{
    const d = await call("/api/chat/"+(c.id==="general"?"general":c.id), {method:"POST", headers:{"Content-Type":"application/json"},
      body: JSON.stringify({message:m, file_content:s.file, sheet_content:s.sheet})});
    addMsg(c.id, "ai", d.reply || "❌ Không có phản hồi", d);
  }catch(e){
    addMsg(c.id, "ai", "❌ Không kết nối được máy chủ (máy chủ miễn phí có thể đang khởi động, hãy thử lại sau ~1 phút).");
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
    } else alert(d.error || "Không tải được tệp");
  }catch(e){ alert("Không tải được tệp: "+e.message); }
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
        else alert(d.error || "Không kết nối được Sheet");
      }catch(e){ alert("Lỗi: "+e.message); }
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
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
