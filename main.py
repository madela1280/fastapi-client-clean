from fastapi import FastAPI, Query, Body, Request, Depends, HTTPException
from sqlalchemy.orm import Session
import requests
import pandas as pd
import os
from datetime import datetime

from cors_config import apply_cors  # ✅ 외부 CORS 적용 함수
from models import Base, Message
from database import engine, SessionLocal

from typing import List
import asyncpg
from pydantic import BaseModel
import json

# ✅ FastAPI 인스턴스는 딱 여기서만 선언
app = FastAPI()
apply_cors(app)  # ✅ CORS 설정도 여기서만 적용

# -------------------------
# 환경 변수 및 캐시
# -------------------------
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")

SHAREPOINT_SITE_ID = "satmoulab.sharepoint.com,102fbb5d-7970-47e4-8686-f6d7fac0375f,cac8f27f-7023-4427-a96f-bd777b42c781"
EXCEL_ITEM_ID = "01BRDK2MMIGCGKWZHSVVEY7CR5K4RRESRZ"
SHEET_NAME = "통합관리"
RANGE_ADDRESS = "H1:Q30000"

_excel_cache = {"data": None, "last_fetched": 0}
CACHE_DURATION = 60

# -------------------------
# 유틸리티 함수들
# -------------------------
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

# -------------------------
# 기본 API
# -------------------------
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

# -------------------------
# PostgreSQL 기반 메시지 저장/조회
# -------------------------
DB_CONFIG = {
    "user": os.environ.get("DB_USER"),
    "password": os.environ.get("DB_PASSWORD"),
    "database": os.environ.get("DB_NAME"),
    "host": os.environ.get("DB_HOST"),
    "port": 5432
}

class MessagePG(BaseModel):
    user_id: str
    role: str
    message: str
    timestamp: str
    read: bool = False

async def get_db_pg():
    return await asyncpg.connect(**DB_CONFIG)

@app.post("/save-message")
async def save_message_pg(msg: MessagePG):
    conn = await get_db_pg()
    await conn.execute("""
        CREATE TABLE IF NOT EXISTS chat_logs (
            id SERIAL PRIMARY KEY,
            user_id TEXT,
            role TEXT,
            message TEXT,
            timestamp TEXT,
            read BOOLEAN
        )
    """)
    await conn.execute("""
        INSERT INTO chat_logs (user_id, role, message, timestamp, read)
        VALUES ($1, $2, $3, $4, $5)
    """, msg.user_id, msg.role, msg.message, msg.timestamp, msg.read)
    await conn.close()
    return {"status": "ok"}

@app.get("/get-messages")
async def get_messages_pg(user_id: str):
    conn = await get_db_pg()
    rows = await conn.fetch(
        "SELECT role, message, timestamp, read FROM chat_logs WHERE user_id=$1 ORDER BY id ASC",
        user_id
    )
    await conn.close()
    return [dict(r) for r in rows]

# -------------------------
# SQLAlchemy 기반 관리자 화면용 메시지 조회
# -------------------------
Base.metadata.create_all(bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

@app.get("/messages/list")
def get_message_list(user_id: str, db: Session = Depends(get_db)):
    messages = db.query(Message).filter(Message.user_id == user_id).order_by(Message.timestamp).all()
    return [
        {
            "sender": m.sender,
            "content": m.content,
            "timestamp": m.timestamp
        }
        for m in messages
    ]

@app.get("/admin/chat-list")
def get_chat_list(db: Session = Depends(get_db)):
    user_ids = db.query(Message.user_id).distinct().all()
    result = []
    for (user_id,) in user_ids:
        latest_msg = (
            db.query(Message)
            .filter(Message.user_id == user_id)
            .order_by(Message.timestamp.desc())
            .first()
        )
        if latest_msg:
            result.append({
                "user_id": latest_msg.user_id,
                "sender": latest_msg.sender,
                "content": latest_msg.content,
                "timestamp": latest_msg.timestamp
            })
    result.sort(key=lambda x: x["timestamp"], reverse=True)
    return result



