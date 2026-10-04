from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
from datetime import datetime
from docx import Document
from openpyxl import Workbook
from xhtml2pdf import pisa
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH ====================
HUGGINGFACE_TOKEN = os.environ.get("HUGGINGFACE_TOKEN", "")
ZALO_BOT_TOKEN = os.environ.get("ZALO_BOT_TOKEN", "")
FB_PAGE_TOKEN = os.environ.get("FB_PAGE_TOKEN", "")
FB_VERIFY_TOKEN = os.environ.get("FB_VERIFY_TOKEN", "baocao_ai_2026")
# ===================================================
