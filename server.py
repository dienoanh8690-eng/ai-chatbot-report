# ==================== CẤU HÌNH — ĐÃ KIỂM TRA CHÍNH XÁC ====================

# Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
# Sửa URL đúng chuẩn mới của Gemini
GEMINI_API_URL_TEMPLATE = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
GEMINI_MODEL = "gemini-2.0-flash-exp"

# Groq / Llama — SỬA URL CHÍNH XÁC
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "llama-3.3-70b-versatile"

# AIML / DOLA
AI_API_KEY = os.environ.get("AI_API_KEY", "").strip()
AI_URL = "https://api.aimlapi.com/v1/chat/completions"
AI_MODEL = os.environ.get("AI_MODEL", "bytedance/dola-seed-2-0-pro")

# OpenAI
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODEL = "gpt-3.5-turbo"

# Claude
CLAUDE_API_KEY = os.environ.get("CLAUDE", "").strip()
CLAUDE_API_URL = "https://api.anthropic.com/v1/messages"

# ==================== HÀM GỌI AI — ĐÃ SỬA LỖI 404 ====================

def goi_gemini(prompt, he_thong=""):
    if not GEMINI_API_KEY:
        return None, "⚠️ Chưa đặt GEMINI_API_KEY"
    url = GEMINI_API_URL_TEMPLATE.format(model=GEMINI_MODEL, key=GEMINI_API_KEY)
    full_text = f"""{he_thong or "Trả lời bằng tiếng Việt rõ ràng."}

Yêu cầu: {prompt}"""
    try:
        res = requests.post(
            url,
            json={"contents": [{"parts": [{"text": full_text}]}]},
            timeout=90
        )
        if res.status_code == 404:
            return None, f"❌ Gemini: Sai model/URL → {GEMINI_MODEL}"
        if res.status_code == 401:
            return None, "❌ Gemini: Khóa không hợp lệ"
        if res.status_code != 200:
            return None, f"❌ Gemini lỗi {res.status_code}: {res.text[:150]}"
        data = res.json()
        if "candidates" not in data:
            return None, f"❌ Gemini không trả lời: {str(data)[:100]}"
        return data["candidates"][0]["content"]["parts"][0]["text"], None
    except Exception as e:
        return None, f"❌ Lỗi kết nối Gemini: {str(e)}"


def goi_groq(prompt, he_thong=""):
    if not GROQ_API_KEY:
        return None, "⚠️ Chưa đặt GROQ_API_KEY"
    try:
        res = requests.post(
            GROQ_API_URL,
            headers={
                "Authorization": f"Bearer {GROQ_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": GROQ_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 4000
            },
            timeout=90
        )
        if res.status_code == 404:
            return None, "❌ Groq: URL hoặc model không tồn tại"
        if res.status_code == 401:
            return None, "❌ Groq: Khóa không hợp lệ"
        if res.status_code != 200:
            return None, f"❌ Groq lỗi {res.status_code}: {res.text[:150]}"
        return res.json()["choices"][0]["message"]["content"], None
    except Exception as e:
        return None, f"❌ Lỗi kết nối Groq: {str(e)}"


def goi_aiml(prompt, he_thong=""):
    if not AI_API_KEY:
        return None, "⚠️ Chưa đặt AI_API_KEY"
    try:
        res = requests.post(
            AI_URL,
            headers={
                "Authorization": f"Bearer {AI_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": AI_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 4000
            },
            timeout=90
        )
        if res.status_code == 404:
            return None, f"❌ AIML: URL/model sai → {AI_MODEL}"
        if res.status_code == 401:
            return None, "❌ AIML: Khóa không hợp lệ"
        if res.status_code != 200:
            return None, f"❌ AIML lỗi {res.status_code}: {res.text[:150]}"
        return res.json()["choices"][0]["message"]["content"], None
    except Exception as e:
        return None, f"❌ Lỗi kết nối AIML: {str(e)}"


def goi_gpt(prompt, he_thong=""):
    if not OPENAI_API_KEY:
        return None, "⚠️ Chưa đặt OPENAI_API_KEY"
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
        if res.status_code == 404:
            return None, "❌ OpenAI: URL/model sai"
        if res.status_code != 200:
            return None, f"❌ GPT lỗi {res.status_code}"
        return res.json()["choices"][0]["message"]["content"], None
    except Exception as e:
        return None, f"❌ Lỗi kết nối GPT: {str(e)}"


def goi_claude(prompt, he_thong=""):
    if not CLAUDE_API_KEY:
        return None, "⚠️ Chưa đặt CLAUDE"
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
                "max_tokens": 4000,
                "system": he_thong or "Trả lời bằng tiếng Việt.",
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=90
        )
        if res.status_code == 404:
            return None, "❌ Claude: URL sai"
        if res.status_code != 200:
            return None, f"❌ Claude lỗi {res.status_code}"
        return res.json()["content"][0]["text"], None
    except Exception as e:
        return None, f"❌ Lỗi kết nối Claude: {str(e)}"
