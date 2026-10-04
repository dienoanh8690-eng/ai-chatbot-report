# ==================================================
# IMPORT — ĐÚNG THỨ TỰ
# ==================================================
from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook
from flask_cors import CORS

# ==================================================
# KHỞI TẠO APP
# ==================================================
app = Flask(__name__)
CORS(app)

# ==================================================
# CẤU HÌNH — ĐÚNG TÊN MODEL ✅
# ==================================================

# AIML / DOLA — ƯU TIÊN SỐ 1
AI_API_KEY = os.environ.get("AI_API_KEY", "").strip()
AI_URL = "https://api.aimlapi.com/v1/chat/completions"
AI_MODEL = "bytedance/dola-seed-2-0-pro"

# Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = "gemini-2.0-flash-exp"
GEMINI_API_URL = f"https://generativelanguage.googleapis.com/v1/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

# Groq / Llama
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "llama-3-1-8b-instant"

# OpenAI / GPT
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODEL = "gpt-3.5-turbo"

# Claude
CLAUDE_API_KEY = os.environ.get("CLAUDE", "").strip()
CLAUDE_API_URL = "https://api.anthropic.com/v1/messages"

# Thư mục
UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# ==================================================
# HÀM GỌI TỪNG AI — TRẢ VỀ (tên, kết quả, lỗi)
# ==================================================
def goi_aiml(prompt, he_thong=""):
    if not AI_API_KEY:
        return "DOLA/AIML", None, "⚠️ Chưa đặt AI_API_KEY"
    try:
        res = requests.post(
            AI_URL,
            headers={"Authorization": f"Bearer {AI_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": AI_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt rõ ràng, tự nhiên."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 3000
            },
            timeout=90
        )
        if res.status_code == 200:
            return "DOLA/AIML", res.json()["choices"][0]["message"]["content"], None
        return "DOLA/AIML", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "DOLA/AIML", None, f"Lỗi kết nối: {str(e)}"


def goi_gemini(prompt, he_thong=""):
    if not GEMINI_API_KEY:
        return "Gemini", None, "⚠️ Chưa đặt GEMINI_API_KEY"
    full_text = f"""{he_thong or "Trả lời bằng tiếng Việt."}

Yêu cầu: {prompt}"""
    try:
        res = requests.post(GEMINI_API_URL, json={"contents": [{"parts": [{"text": full_text}]}]}, timeout=90)
        if res.status_code == 200:
            data = res.json()
            if "candidates" in data:
                return "Gemini", data["candidates"][0]["content"]["parts"][0]["text"], None
        return "Gemini", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "Gemini", None, f"Lỗi kết nối: {str(e)}"


def goi_groq(prompt, he_thong=""):
    if not GROQ_API_KEY:
        return "Llama/Groq", None, "⚠️ Chưa đặt GROQ_API_KEY"
    try:
        res = requests.post(
            GROQ_API_URL,
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": GROQ_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 3000
            },
            timeout=90
        )
        if res.status_code == 200:
            return "Llama/Groq", res.json()["choices"][0]["message"]["content"], None
        return "Llama/Groq", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "Llama/Groq", None, f"Lỗi kết nối: {str(e)}"


def goi_gpt(prompt, he_thong=""):
    if not OPENAI_API_KEY:
        return "GPT", None, "⚠️ Chưa đặt OPENAI_API_KEY"
    try:
        res = requests.post(
            OPENAI_API_URL,
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": OPENAI_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7
            },
            timeout=90
        )
        if res.status_code == 200:
            return "GPT", res.json()["choices"][0]["message"]["content"], None
        return "GPT", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "GPT", None, f"Lỗi kết nối: {str(e)}"


def goi_claude(prompt, he_thong=""):
    if not CLAUDE_API_KEY:
        return "Claude", None, "⚠️ Chưa đặt CLAUDE"
    try:
        res = requests.post(
            CLAUDE_API_URL,
            headers={
                "x-api-key": CLAUDE_API_KEY,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json"
            },
            json={
                "model": "claude-3-5-sonnet-20241022",
                "max_tokens": 3000,
                "system": he_thong or "Trả lời bằng tiếng Việt.",
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=90
        )
        if res.status_code == 200:
            return "Claude", res.json()["content"][0]["text"], None
        return "Claude", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "Claude", None, f"Lỗi kết nối: {str(e)}"


# ==================================================
# LUỒNG GỌI — ĐÚNG THỨ TỰ TỪNG MỤC ✅
# ==================================================
def goi_theo_danh_sach(prompt, danh_sach_ham, he_thong=""):
    """Chạy tuần tự, trả về AI đầu tiên thành công"""
    loi_tong = []
    for ham in danh_sach_ham:
        ten, kq, loi = ham(prompt, he_thong=he_thong)
        if kq:
            return f"✅ [{ten}]\n{kq}"
        loi_tong.append(f"{ten}: {loi}")
    return "❌ Tất cả AI đều không trả lời:\n" + "\n".join(f"× {x}" for x in loi_tong)


# ==================================================
# ĐỌC GOOGLE SHEETS & TẠO TỆP
# ==================================================
def doc_google_sheet(sheet_url):
    try:
        if "docs.google.com/spreadsheets/d/" not in sheet_url:
            return None, "❌ Link không đúng định dạng Google Sheets"
        sheet_id = sheet_url.split("/d/")[1].split("/")[0]
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        res = requests.get(csv_url, timeout=30)
        if res.status_code == 200:
            return res.text, None
        return None, "❌ Không đọc được Sheet → Kiểm tra quyền chia sẻ: Bất kỳ ai có link"
    except Exception as e:
        return None, f"❌ Lỗi đọc Sheet: {str(e)}"


def tao_word(noi_dung):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
        duong_dan = os.path.join(RESULT_FOLDER, ten)
        doc = Document()
        p = doc.add_heading("BÁO CÁO", 0)
        for run in p.runs:
            run.font.name = "Arial"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        doc.add_paragraph("-" * 60)
        for dong in noi_dung.split("\n"):
            if dong.strip():
                p = doc.add_paragraph(dong.strip())
                for run in p.runs:
                    run.font.name = "Arial"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        doc.save(duong_dan)
        return ten
    except:
        return ""


def tao_excel(noi_dung=""):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
        duong_dan = os.path.join(RESULT_FOLDER, ten)
        wb = Workbook()
        ws = wb.active
        ws.title = "BÁO CÁO"
        ws["A1"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        hang = 3
        for dong in noi_dung.split("\n"):
            if dong.strip():
                ws.cell(row=hang, column=1, value=dong.strip())
                hang += 1
        wb.save(duong_dan)
        return ten
    except:
        return ""


# ==================================================
# ROUTE API — ĐÚNG AI CHO TỪNG MỤC ✅
# ==================================================

@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Không có tệp"}), 400
    f = request.files["file"]
    ext = f.filename.rsplit(".", 1)[-1].lower()
    ten_moi = f"{uuid.uuid4().hex[:10]}.{ext}"
    duong_dan = os.path.join(UPLOAD_FOLDER, ten_moi)
    f.save(duong_dan)
    noi_dung = ""
    try:
        if ext == "xlsx":
            wb = load_workbook(duong_dan, data_only=True, read_only=True)
            ws = wb.active
            for hang in ws.iter_rows(values_only=True):
                noi_dung += " | ".join(str(c) if c is not None else "" for c in hang) + "\n"
            wb.close()
        elif ext == "txt":
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
    except Exception as e:
        noi_dung = f"(Không đọc được nội dung: {e})"
    return jsonify({"status": "ok", "name": f.filename, "content": noi_dung[:5000]})


@app.route("/api/connect-sheet", methods=["POST"])
def connect_sheet():
    data = request.get_json(silent=True) or {}
    url = data.get("url", "").strip()
    if not url:
        return jsonify({"error": "Vui lòng dán link Google Sheets"}), 400
    noi_dung, loi = doc_google_sheet(url)
    if loi:
        return jsonify({"error": loi}), 400
    return jsonify({"status": "ok", "content": noi_dung[:5000]})


@app.route("/api/chat-data", methods=["POST"])
def chat_data():
    """📊 Xử lý dữ liệu — DOLA/AIML → Gemini → Llama/Groq"""
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    sheet_content = data.get("sheet_content", "")
    full_data = f"{file_content}\n---\n{sheet_content}".strip()
    if not msg and not full_data:
        return jsonify({"reply": "Vui lòng nhập yêu cầu, tải tệp hoặc dán link Google Sheets!"})
    prompt = f"{msg}\n\nDữ liệu phân tích:\n{full_data}" if full_data else msg
    he_thong = "Bạn là chuyên gia phân tích dữ liệu và tạo báo cáo. Trả lời rõ ràng, tóm tắt số liệu quan trọng, dùng bảng khi phù hợp."
    
    danh_sach = [goi_aiml, goi_gemini, goi_groq]  # ✅ KHÔNG CÓ CLAUDE
    tra_loi = goi_theo_danh_sach(prompt, danh_sach, he_thong)
    
    word = tao_word(tra_loi) if "✅" in tra_loi else ""
    excel = tao_excel(tra_loi) if "✅" in tra_loi else ""
    return jsonify({
        "reply": tra_loi,
        "word": f"/download/{word}" if word else "",
        "excel": f"/download/{excel}" if excel else ""
    })


@app.route("/api/chat-doc", methods=["POST"])
def chat_doc():
    """✍️ Soạn thảo — DOLA/AIML → GPT → Claude"""
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    if not msg:
        return jsonify({"reply": "Vui lòng nhập yêu cầu soạn thảo!"})
    he_thong = "Bạn là chuyên gia soạn thảo văn bản hành chính, hợp đồng, thư từ. Viết chuẩn mực, đúng thể thức Việt Nam."
    
    danh_sach = [goi_aiml, goi_gpt, goi_claude]  # ✅ ĐÚNG: có Claude ở cuối
    tra_loi = goi_theo_danh_sach(msg, danh_sach, he_thong)
    return jsonify({"reply": tra_loi})


@app.route("/api/chat-tender", methods=["POST"])
def chat_tender():
    """🏆 Đấu thầu — DOLA/AIML → Gemini → Llama/Groq"""
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    if not msg:
        return jsonify({"reply": "Vui lòng nhập yêu cầu về quy trình đấu thầu!"})
    he_thong = "Bạn là chuyên gia tư vấn quy trình đấu thầu theo pháp luật Việt Nam. Hướng dẫn chi tiết từng bước, hồ sơ, lưu ý pháp lý."
    
    danh_sach = [goi_aiml, goi_gemini, goi_groq]  # ✅ KHÔNG CÓ CLAUDE
    tra_loi = goi_theo_danh_sach(msg, danh_sach, he_thong)
    return jsonify({"reply": tra_loi})


@app.route("/api/chat-equip", methods=["POST"])
def chat_equip():
    """🔧 Thiết bị — DOLA/AIML → Gemini → Llama/Groq"""
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    full_prompt = f"{msg}\n\nDữ liệu thiết bị:\n{file_content[:3000]}" if file_content else msg
    if not msg and not file_content:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải danh sách thiết bị!"})
    he_thong = "Bạn là chuyên gia quản lý thiết bị nhà máy. Phân loại, theo dõi tình trạng, đề xuất bảo trì, tính tuổi thọ."
    
    danh_sach = [goi_aiml, goi_gemini, goi_groq]  # ✅ KHÔNG CÓ CLAUDE
    tra_loi = goi_theo_danh_sach(full_prompt, danh_sach, he_thong)
    return jsonify({"reply": tra_loi})


@app.route("/download/<ten_file>")
def download(ten_file):
    for folder in [RESULT_FOLDER, UPLOAD_FOLDER]:
        path = os.path.join(folder, ten_file)
        if os.path.exists(path):
            return send_file(path, as_attachment=True)
    return "Không tìm thấy tệp", 404


@app.route("/")
def trang_chu():
    return r"""
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Thủy Điện — Hệ thống hỗ trợ toàn diện</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
--p1: #2563eb; --p1-light: #dbeafe; --p1-bg: #eff6ff;
--p2: #16a34a; --p2-light: #dcfce7; --p2-bg: #f0fdf4;
--p3: #9333ea; --p3-light: #f3e8ff; --p3-bg: #faf5ff;
--p4: #f59e0b; --p4-light: #fef3c7; --p4-bg: #fffbeb;
--gray-100: #f1f5f9; --gray-200: #e2e8f0; --gray-600: #475569; --gray-800: #1e293b;
--shadow-sm: 0 1px 3px rgba(0,0,0,0.05);
--shadow-md: 0 4px 12px rgba(0,0,0,0.06);
--shadow-lg: 0 10px 30px rgba(37,99,235,0.08);
--radius-sm: 8px; --radius-md: 12px; --radius-lg: 20px;
}
* { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Inter', sans-serif; }
body { background: linear-gradient(135deg, #f0f7ff 0%, #faf5ff 100%); min-height: 100vh; padding: 20px; }
.header { text-align: center; margin-bottom: 28px; padding-top: 10px; }
.header h1 { font-size: 26px; font-weight: 700; color: var(--gray-800); margin-bottom: 6px; }
.header p { color: var(--gray-600); font-size: 14px; }
.grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 22px; max-width: 1920px; margin: 0 auto; }
@media (max-width: 1400px) { .grid { grid-template-columns: repeat(2, 1fr); } }
@media (max-width: 768px) { .grid { grid-template-columns: 1fr; } }
.card { background: white; border-radius: var(--radius-lg); padding: 24px; box-shadow: var(--shadow-md); display: flex; flex-direction: column; height: calc(100vh - 160px); min-height: 650px; transition: transform 0.2s, box-shadow 0.2s; }
.card:hover { transform: translateY(-2px); box-shadow: var(--shadow-lg); }
.card-p1 { border-top: 4px solid var(--p1); }
.card-p2 { border-top: 4px solid var(--p2); }
.card-p3 { border-top: 4px solid var(--p3); }
.card-p4 { border-top: 4px solid var(--p4); }
.card-head { display: flex; align-items: center; gap: 10px; margin-bottom: 18px; padding-bottom: 14px; border-bottom: 2px solid; }
.card-p1 .card-head { border-bottom-color: var(--p1-light); }
.card-p2 .card-head { border-bottom-color: var(--p2-light); }
.card-p3 .card-head { border-bottom-color: var(--p3-light); }
.card-p4 .card-head { border-bottom-color: var(--p4-light); }
.card-icon { font-size: 24px; }
.card-title { font-size: 17px; font-weight: 700; }
.card-p1 .card-title { color: var(--p1); }
.card-p2 .card-title { color: var(--p2); }
.card-p3 .card-title { color: var(--p3); }
.card-p4 .card-title { color: var(--p4); }
.card-ai { font-size: 11px; color: #94a3b8; margin-left: auto; background: var(--gray-100); padding: 3px 8px; border-radius: 12px; }
.upload-zone { border: 2px dashed var(--gray-200); border-radius: var(--radius-md); padding: 18px; text-align: center; cursor: pointer; margin-bottom: 12px; transition: all 0.25s; }
.card-p1 .upload-zone:hover { border-color: var(--p1); background: var(--p1-bg); }
.card-p4 .upload-zone:hover { border-color: var(--p4); background: var(--p4-bg); }
.upload-zone p { font-size: 13px; color: var(--gray-600); }
.file-bar { display: flex; align-items: center; gap: 8px; padding: 10px 14px; border-radius: var(--radius-sm); margin-bottom: 12px; font-size: 13px; font-weight: 500; }
.card-p1 .file-bar { background: var(--p1-light); color: #1e40af; }
.card-p4 .file-bar { background: var(--p4-light); color: #92400e; }
.file-bar button { margin-left: auto; background: none; border: none; font-size: 18px; cursor: pointer; color: #dc2626; line-height: 1; }
.sheet-bar { display: flex; gap: 8px; margin-bottom: 12px; }
.sheet-bar input { flex: 1; padding: 10px 14px; border: 1px solid var(--gray-200); border-radius: var(--radius-sm); font-size: 13px; outline: none; transition: border-color 0.2s; }
.sheet-bar input:focus { border-color: var(--p1); }
.sheet-bar button { padding: 10px 16px; border: none; border-radius: var(--radius-sm); background: var(--p1); color: white; cursor: pointer; font-weight: 600; font-size: 13px; transition: background 0.2s; }
.sheet-bar button:hover { background: #1d4ed8; }
.quick-btns { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 14px; }
.q-btn { padding: 10px 12px; border: 1px solid var(--gray-200); border-radius: var(--radius-sm); background: white; cursor: pointer; font-size: 12px; transition: all 0.2s; text-align: left; line-height: 1.4; }
.card-p1 .q-btn:hover { border-color: var(--p1); background: var(--p1-bg); transform: translateY(-1px); }
.card-p2 .q-btn:hover { border-color: var(--p2); background: var(--p2-bg); transform: translateY(-1px); }
.card-p3 .q-btn:hover { border-color: var(--p3); background: var(--p3-bg); transform: translateY(-1px); }
.card-p4 .q-btn:hover { border-color: var(--p4); background: var(--p4-bg); transform: translateY(-1px); }
.chat-area { flex: 1; overflow-y: auto; padding: 4px; margin-bottom: 14px; }
.msg { margin-bottom: 16px; max-width: 98%; animation: fadeIn 0.3s forwards; opacity: 0; }
@keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
.msg.user { margin-left: auto; }
.msg.ai { margin-right: auto; }
.bubble { padding: 14px 18px; border-radius: 18px; line-height: 1.6; font-size: 14px; white-space: pre-wrap; word-break: break-word; }
.msg.user .bubble { background: linear-gradient(135deg, #dbeafe, #e0e7ff); border-bottom-right-radius: 8px; color: #1e3a8a; }
.msg.ai .bubble { background: var(--gray-100); border-bottom-left-radius: 8px; color: var(--gray-800); }
.bubble.warn { background: #fef3c7; border-left: 3px solid #f59e0b; color: #92400e; }
.bubble.err { background: #fee2e2; border-left: 3px solid #ef4444; color: #b91c1c; }
.dl-group { display: flex; gap: 10px; margin-top: 12px; flex-wrap: wrap; }
.dl-btn { display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px; border-radius: 20px; text-decoration: none; font-size: 13px; font-weight: 600; transition: transform 0.2s; }
.dl-btn:hover { transform: scale(1.05); }
.dl-word { background: var(--p1-light); color: #1d4ed8; }
.dl-excel { background: var(--p2-light); color: #15803d; }
.input-row { display: flex; gap: 10px; align-items: flex-end; }
textarea { flex: 1; min-height: 48px; max-height: 120px; padding: 12px 18px; border: 1px solid var(--gray-200); border-radius: 24px; font-size: 14px; resize: none; outline: none; transition: all 0.2s; line-height: 1.5; }
textarea:focus { border-color: var(--p1); box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
.card-p2 textarea:focus { border-color: var(--p2); box-shadow: 0 0 0 3px rgba(22,163,74,0.1); }
.card-p3 textarea:focus { border-color: var(--p3); box-shadow: 0 0 0 3px rgba(147,51,234,0.1); }
.card-p4 textarea:focus { border-color: var(--p4); box-shadow: 0 0 0 3px rgba(245,158,11,0.1); }
.send-btn { width: 44px; height: 44px; border-radius: 50%; border: none; color: white; cursor: pointer; font-size: 18px; flex-shrink: 0; display: flex; align-items: center; justify-content: center; transition: all 0.2s; }
.send-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
.card-p1 .send-btn { background: var(--p1); }
.card-p2 .send-btn { background: var(--p2); }
.card-p3 .send-btn { background: var(--p3); }
.card-p4 .send-btn { background: var(--p4); }
.send-btn:hover:not(:disabled) { transform: scale(1.1); }
.hidden { display: none !important; }
</style>
</head>
<body>
<div class="header">
<h1>⚡ All Thủy Điện — Hệ thống hỗ trợ toàn diện</h1>
<p>Phân tích dữ liệu · Soạn thảo văn bản · Tư vấn đấu thầu · Quản lý thiết bị</p>
</div>
<div class="grid">

<!-- CỘT 1: XỬ LÝ DỮ LIỆU & TẠO BÁO CÁO -->
<div class="card card-p1">
<div class="card-head">
<span class="card-icon">📊</span>
<h3 class="card-title">Xử lý dữ liệu & Tạo báo cáo</h3>
<span class="card-ai">DOLA/AIML → Gemini → Llama</span>
</div>

<div class="upload-zone" id="uploadZone1" onclick="document.getElementById('fileInput1').click()">
<p>📎 Nhấn chọn hoặc kéo thả tệp (.xlsx, .txt)</p>
</div>
<input type="file" id="fileInput1" accept=".xlsx,.txt" class="hidden">
<div class="file-bar hidden" id="fileBar1">
<span id="fName1"></span>
<button onclick="clearFile1()">✕</button>
</div>

<div class="sheet-bar">
<input type="text" id="sheetUrl" placeholder="🔗 Dán link Google Sheets...">
<button onclick="connectSheet()">Kết nối</button>
</div>
<div class="file-bar hidden" id="sheetBar">
<span>✅ Google Sheets đã kết nối</span>
<button onclick="clearSheet()">✕</button>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickData('Sắp xếp và tóm tắt dữ liệu')">📋 Sắp xếp dữ liệu</button>
<button class="q-btn" onclick="quickData('Tính tổng và phân tích số liệu')">💰 Tính tổng hợp</button>
<button class="q-btn" onclick="quickData('Lập báo cáo đầy đủ có cấu trúc')">📑 Lập báo cáo</button>
<button class="q-btn" onclick="quickData('Đánh giá xu hướng và đề xuất')">📈 Nhận xét & Đề xuất</button>
</div>

<div class="chat-area" id="chat1">
<div class="msg ai"><div class="bubble">👋 Tải tệp, dán link Google Sheets hoặc nhập yêu cầu để bắt đầu phân tích nhé!</div></div>
</div>

<div class="input-row">
<textarea id="input1" placeholder="Nhập yêu cầu phân tích..." onkeydown="handleKey(event, sendData)"></textarea>
<button class="send-btn" id="btn1" onclick="sendData()">➤</button>
</div>
</div>

<!-- CỘT 2: SOẠN THẢO VĂN BẢN -->
<div class="card card-p2">
<div class="card-head">
<span class="card-icon">✍️</span>
<h3 class="card-title">Soạn thảo văn bản</h3>
<span class="card-ai">DOLA/AIML → GPT → Claude</span>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickDoc('Soạn thảo công văn gửi cấp trên')">📝 Công văn</button>
<button class="q-btn" onclick="quickDoc('Soạn thảo hợp đồng mua bán thiết bị')">📄 Hợp đồng</button>
<button class="q-btn" onclick="quickDoc('Viết báo cáo tiến độ thực hiện dự án')">📈 Báo cáo tiến độ</button>
<button class="q-btn" onclick="quickDoc('Soạn thảo thư mời họp và biên bản')">📋 Thư & Biên bản</button>
</div>

<div class="chat-area" id="chat2">
<div class="msg ai"><div class="bubble">👋 Tôi sẽ giúp bạn soạn thảo văn bản chuẩn mực, đúng thể thức Việt Nam. Bạn cần viết gì?</div></div>
</div>

<div class="input-row">
<textarea id="input2" placeholder="Bạn cần soạn thảo gì...?" onkeydown="handleKey(event, sendDoc)"></textarea>
<button class="send-btn" id="btn2" onclick="sendDoc()">➤</button>
</div>
</div>

<!-- CỘT 3: QUY TRÌNH ĐẤU THẦU -->
<div class="card card-p3">
<div class="card-head">
<span class="card-icon">🏆</span>
<h3 class="card-title">Quy trình đấu thầu</h3>
<span class="card-ai">DOLA/AIML → Gemini → Llama</span>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickTender('Giải thích toàn bộ quy trình đấu thầu')">📋 Toàn bộ quy trình</button>
<button class="q-btn" onclick="quickTender('Danh mục hồ sơ cần chuẩn bị')">📑 Hồ sơ mời thầu</button>
<button class="q-btn" onclick="quickTender('Lưu ý pháp lý và rủi ro thường gặp')">⚖️ Pháp lý & Rủi ro</button>
<button class="q-btn" onclick="quickTender('Mẫu biểu mẫu thông dụng')">📄 Biểu mẫu</button>
</div>

<div class="chat-area" id="chat3">
<div class="msg ai"><div class="bubble">👋 Tôi hướng dẫn chi tiết theo quy định Việt Nam. Bạn cần hỗ trợ về bước nào?</div></div>
</div>

<div class="input-row">
<textarea id="input3" placeholder="Hỏi về quy trình đấu thầu...?" onkeydown="handleKey(event, sendTender)"></textarea>
<button class="send-btn" id="btn3" onclick="sendTender()">➤</button>
</div>
</div>

<!-- CỘT 4: QUẢN LÝ THIẾT BỊ -->
<div class="card card-p4">
<div class="card-head">
<span class="card-icon">🔧</span>
<h3 class="card-title">Quản lý thiết bị</h3>
<span class="card-ai">DOLA/AIML → Gemini → Llama</span>
</div>

<div class="upload-zone" id="uploadZone4" onclick="document.getElementById('fileInput4').click()">
<p>📎 Tải danh sách thiết bị (.xlsx, .txt)</p>
</div>
<input type="file" id="fileInput4" accept=".xlsx,.txt" class="hidden">
<div class="file-bar hidden" id="fileBar4">
<span id="fName4"></span>
<button onclick="clearFile4()">✕</button>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickEquip('Phân loại thiết bị theo nhóm')">📊 Phân loại thiết bị</button>
<button class="q-btn" onclick="quickEquip('Đề xuất kế hoạch bảo trì định kỳ')">🛠️ Kế hoạch bảo trì</button>
<button class="q-btn" onclick="quickEquip('Đánh giá tình trạng và rủi ro')">⚠️ Đánh giá rủi ro</button>
<button class="q-btn" onclick="quickEquip('Tính tuổi thọ và đề xuất thay thế')">🔄 Tuổi thọ & Thay thế</button>
</div>

<div class="chat-area" id="chat4">
<div class="msg ai"><div class="bubble">👋 Tải danh sách thiết bị hoặc nhập yêu cầu — tôi sẽ phân tích chi tiết nhé!</div></div>
</div>

<div class="input-row">
<textarea id="input4" placeholder="Nhập yêu cầu quản lý thiết bị...?" onkeydown="handleKey(event, sendEquip)"></textarea>
<button class="send-btn" id="btn4" onclick="sendEquip()">➤</button>
</div>
</div>

</div>

<script>
let file1Content = "";
let file4Content = "";
let sheetContent = "";

document.getElementById('fileInput1').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    const fd = new FormData(); fd.append('file', f);
    fetch('/api/upload', {method:'POST', body:fd}).then(r=>r.json()).then(d=>{
        if(d.status==='ok'){
            file1Content = d.content;
            document.getElementById('fName1').textContent = d.name;
            document.getElementById('fileBar1').classList.remove('hidden');
        }
    });
});
document.getElementById('fileInput4').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    const fd = new FormData(); fd.append('file', f);
    fetch('/api/upload', {method:'POST', body:fd}).then(r=>r.json()).then(d=>{
        if(d.status==='ok'){
            file4Content = d.content;
            document.getElementById('fName4').textContent = d.name;
            document.getElementById('fileBar4').classList.remove('hidden');
        }
    });
});
function clearFile1(){ file1Content=''; document.getElementById('fileBar1').classList.add('hidden'); document.getElementById('fileInput1').value=''; }
function clearFile4(){ file4Content=''; document.getElementById('fileBar4').classList.add('hidden'); document.getElementById('fileInput4').value=''; }
function clearSheet(){ sheetContent=''; document.getElementById('sheetBar').classList.add('hidden'); document.getElementById('sheetUrl').value=''; }

function connectSheet(){
    const url = document.getElementById('sheetUrl').value.trim();
    if(!url) return;
    fetch('/api/connect-sheet', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({url})})
    .then(r=>r.json()).then(d=>{
        if(d.status==='ok'){
            sheetContent = d.content;
            document.getElementById('sheetBar').classList.remove('hidden');
        } else alert(d.error);
    });
}

function handleKey(e, fn){ if(e.key==='Enter' && !e.shiftKey){ e.preventDefault(); fn(); } }
function addMsg(chatId, type, text, links=''){
    const c = document.getElementById(chatId);
    const d = document.createElement('div'); d.className = 'msg '+type;
    let cls = '';
    if(text.includes('⚠️')) cls=' warn';
    else if(text.includes('❌')) cls=' err';
    d.innerHTML = `<div class="bubble${cls}">${text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}${links}</div>`;
    c.appendChild(d); c.scrollTop = c.scrollHeight;
}
function dlLinks(d){
    let h='';
    if(d.word) h += `<a href="${d.word}" class="dl-btn dl-word" target="_blank">📄 Tải Word</a>`;
    if(d.excel) h += `<a href="${d.excel}" class="dl-btn dl-excel" target="_blank">📊 Tải Excel</a>`;
    return h ? `<div class="dl-group">${h}</div>` : '';
}

async function sendData(){
    const i = document.getElementById('input1');
    const m = i.value.trim(); if(!m && !file1Content && !sheetContent) return;
    addMsg('chat1','user',m||'Phân tích dữ liệu'); i.value='';
    document.getElementById('btn1').disabled=true; document.getElementById('btn1').textContent='⏳';
    const r = await fetch('/api/chat-data', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m, file_content:file1Content, sheet_content:sheetContent})});
    const d = await r.json();
    addMsg('chat1','ai',d.reply, dlLinks(d));
    document.getElementById('btn1').disabled=false; document.getElementById('btn1').textContent='➤';
}
function quickData(t){ document.getElementById('input1').value=t; sendData(); }

async function sendDoc(){
    const i = document.getElementById('input2');
    const m = i.value.trim(); if(!m) return;
    addMsg('chat2','user',m); i.value='';
    document.getElementById('btn2').disabled=true; document.getElementById('btn2').textContent='⏳';
    const r = await fetch('/api/chat-doc', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m})});
    const d = await r.json(); addMsg('chat2','ai',d.reply);
    document.getElementById('btn2').disabled=false; document.getElementById('btn2').textContent='➤';
}
function quickDoc(t){ document.getElementById('input2').value=t; sendDoc(); }

async function sendTender(){
    const i = document.getElementById('input3');
    const m = i.value.trim(); if(!m) return;
    addMsg('chat3','user',m); i.value='';
    document.getElementById('btn3').disabled=true; document.getElementById('btn3').textContent='⏳';
    const r = await fetch('/api/chat-tender', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m})});
    const d = await r.json(); addMsg('chat3','ai',d.reply);
    document.getElementById('btn3').disabled=false; document.getElementById('btn3').textContent='➤';
}
function quickTender(t){ document.getElementById('input3').value=t; sendTender(); }

async function sendEquip(){
    const i = document.getElementById('input4');
    const m = i.value.trim(); if(!m && !file4Content) return;
    addMsg('chat4','user',m||'Phân tích danh sách thiết bị'); i.value='';
    document.getElementById('btn4').disabled=true; document.getElementById('btn4').textContent='⏳';
    const r = await fetch('/api/chat-equip', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m, file_content:file4Content})});
    const d = await r.json(); addMsg('chat4','ai',d.reply);
    document.getElementById('btn4').disabled=false; document.getElementById('btn4').textContent='➤';
}
function quickEquip(t){ document.getElementById('input4').value=t; sendEquip(); }
</script>
</body>
</html>
"""

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
