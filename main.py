from fastapi import FastAPI, Query, Body, Request
from fastapi.middleware.cors import CORSMiddleware
import requests
import pandas as pd
import os
from datetime import datetime

app = FastAPI()

# ✅ CORS 정확히 허용할 Netlify 도메인 명시
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://storied-kitsune-a986bd.netlify.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 환경 변수
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")

SHAREPOINT_SITE_ID = "satmoulab.sharepoint.com,102fbb5d-7970-47e4-8686-f6d7fac0375f,cac8f27f-7023-4427-a96f-bd777b42c781"
EXCEL_ITEM_ID = "01BRDK2MMIGCGKWZHSVVEY7CR5K4RRESRZ"
SHEET_NAME = "통합관리"
RANGE_ADDRESS = "H1:Q30000"

# 캐시 (선택 사항 - 속도 개선용)
_excel_cache = {"data": None, "last_fetched": 0}
CACHE_DURATION = 60

# 헬퍼 함수들
def normalize_phone(p):
    return str(p).replace("-", "").replace(" ", "").strip()

def parse_excel_date(value):
    if isinstance(value, (float, int)):
        base_date = datetime(1899, 12, 30)
        return (base_date + pd.to_timedelta(value, unit="D")).strftime("%Y-%m-%d")
    if isinstance(value, str):
        try:
            return datetime.strptime(value[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        except:
            return value
    return str(value)

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    response = requests.post(url, headers=headers, data=data)
    return response.json()["access_token"]

def get_excel_data(phone: str):
    token = get_access_token()
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{EXCEL_ITEM_ID}/workbook/worksheets('{SHEET_NAME}')/range(address='{RANGE_ADDRESS}')"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    data = response.json()

    values = data.get("values")
    if not values:
        return None

    header = values[0]
    rows = values[1:]

    try:
        phone = normalize_phone(phone)
        contact1_idx = header.index("연락처1")
        contact2_idx = header.index("연락처2")
        name_idx = header.index("수취인명")
        start_idx = header.index("시작일")
        end_idx = header.index("종료일")
        model_idx = header.index("제품명")
        return_idx = header.index("반납완료일") if "반납완료일" in header else None
    except ValueError:
        return None

    for row in reversed(rows):
        contact1 = normalize_phone(row[contact1_idx]) if contact1_idx < len(row) else ""
        contact2 = normalize_phone(row[contact2_idx]) if contact2_idx < len(row) else ""
        is_returned = row[return_idx] if return_idx is not None and len(row) > return_idx else None

        if phone == contact1 or phone == contact2:
            if not is_returned:
                return {
                    "대여자명": row[name_idx],
                    "대여시작일": parse_excel_date(row[start_idx]),
                    "대여종료일": parse_excel_date(row[end_idx]),
                    "제품명": row[model_idx] if model_idx < len(row) else ""
                }

    return {
        "대여자명": None,
        "대여시작일": None,
        "대여종료일": None,
        "제품명": None
    }

# 📌 입금 문자 수신 저장용
deposit_logs = []

@app.get("/")
def root():
    return {"message": "FastAPI Excel 연결 OK"}

@app.get("/get-user-info")
def get_user_info(phone: str = Query(...)):
    return get_excel_data(phone)

@app.post("/deposit-webhook")
async def handle_sms(request: Request):
    content_type = request.headers.get("content-type", "")

    if "application/json" in content_type:
        body = await request.json()
    elif "application/x-www-form-urlencoded" in content_type:
        form = await request.form()
        body = dict(form)
    else:
        return {"error": "Unsupported content-type"}

    deposit_logs.append(body)
    return {"status": "received"}

@app.get("/deposit-log")
def get_deposit_logs():
    return deposit_logs

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)

# --- 이 코드를 기존 main.py 제일 아래에 추가하세요 ---

from pydantic import BaseModel
from typing import List
import asyncpg
import json

# DB 연결 정보 (Render PostgreSQL 기준)
DB_CONFIG = {
    "user": os.environ.get("DB_USER"),
    "password": os.environ.get("DB_PASSWORD"),
    "database": os.environ.get("DB_NAME"),
    "host": os.environ.get("DB_HOST"),
    "port": 5432
}

# 메시지 모델 정의
class Message(BaseModel):
    user_id: str
    role: str  # 'user' or 'bot'
    message: str
    timestamp: str  # ISO8601 문자열
    read: bool = False

# PostgreSQL 연결
async def get_db():
    return await asyncpg.connect(**DB_CONFIG)

# 메시지 저장 API
@app.post("/save-message")
async def save_message(msg: Message):
    conn = await get_db()
    await conn.execute(
        """
        CREATE TABLE IF NOT EXISTS chat_logs (
            id SERIAL PRIMARY KEY,
            user_id TEXT,
            role TEXT,
            message TEXT,
            timestamp TEXT,
            read BOOLEAN
        )
        """
    )
    await conn.execute(
        """
        INSERT INTO chat_logs (user_id, role, message, timestamp, read)
        VALUES ($1, $2, $3, $4, $5)
        """,
        msg.user_id, msg.role, msg.message, msg.timestamp, msg.read
    )
    await conn.close()
    return {"status": "ok"}

# 메시지 불러오기 API
@app.get("/get-messages")
async def get_messages(user_id: str):
    conn = await get_db()
    rows = await conn.fetch(
        "SELECT role, message, timestamp, read FROM chat_logs WHERE user_id=$1 ORDER BY id ASC",
        user_id
    )
    await conn.close()
    return [dict(r) for r in rows]

# 메시지 삭제 API + 백업용 엑셀 저장은 다음 단계에서




