from fastapi import FastAPI, Query, Body, Request, Depends
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
import requests
import pandas as pd
import os
from datetime import datetime
from models import Base, Message, MessageCreate  # ✅ 이 줄이 중요!
from database import engine, SessionLocal
from typing import List
import json

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

_excel_cache = {"data": None, "last_fetched": 0}
CACHE_DURATION = 60

CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")
SHAREPOINT_SITE_ID = "satmoulab.sharepoint.com,102fbb5d-7970-47e4-8686-f6d7fac0375f,cac8f27f-7023-4427-a96f-bd777b42c781"
EXCEL_ITEM_ID = "01BRDK2MMIGCGKWZHSVVEY7CR5K4RRESRZ"
SHEET_NAME = "통합관리"
RANGE_ADDRESS = "H1:Q30000"

def normalize_phone(p): return str(p).replace("-", "").replace(" ", "").strip()

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
        contact1_idx = header.index("\uc5f0\ub78c\ucc981")
        contact2_idx = header.index("\uc5f0\ub78c\ucc982")
        name_idx = header.index("\uc218\uce58\uc778\uba85")
        start_idx = header.index("\uc2dc\uc791\uc77c")
        end_idx = header.index("\uc885\ub8cc\uc77c")
        model_idx = header.index("\uc81c\ud488\uba85")
        return_idx = header.index("\ubc18\ub0a9\uc644\ub8cc\uc77c") if "\ubc18\ub0a9\uc644\ub8cc\uc77c" in header else None
    except ValueError:
        return None

    for row in reversed(rows):
        contact1 = normalize_phone(row[contact1_idx]) if contact1_idx < len(row) else ""
        contact2 = normalize_phone(row[contact2_idx]) if contact2_idx < len(row) else ""
        is_returned = row[return_idx] if return_idx is not None and len(row) > return_idx else None
        if phone == contact1 or phone == contact2:
            if not is_returned:
                return {
                    "\ub300\uc5ec\uc790\uba85": row[name_idx],
                    "\ub300\uc5ec\uc2dc\uc791\uc77c": parse_excel_date(row[start_idx]),
                    "\ub300\uc5ec\uc885\ub8cc\uc77c": parse_excel_date(row[end_idx]),
                    "\uc81c\ud488\uba85": row[model_idx] if model_idx < len(row) else ""
                }

    return {"\ub300\uc5ec\uc790\uba85": None, "\ub300\uc5ec\uc2dc\uc791\uc77c": None, "\ub300\uc5ec\uc885\ub8cc\uc77c": None, "\uc81c\ud488\uba85": None}

@app.get("/")
def root():
    return {"message": "FastAPI Excel \uc5f0\uacb0 OK"}

@app.get("/get-user-info")
def get_user_info(phone: str = Query(...)):
    return get_excel_data(phone)

deposit_logs = []

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

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

@app.post("/save-message")
async def save_message(msg: MessageCreate, db: Session = Depends(get_db)):
    new_msg = Message(
        user_id=msg.user_id,
        sender=msg.role,
        content=msg.message,
        timestamp=msg.timestamp,
        read=msg.read
    )
    db.add(new_msg)
    db.commit()
    return {"status": "ok"}

@app.get("/messages/list")
def get_message_list(user_id: str, db: Session = Depends(get_db)):
    messages = db.query(Message).filter(
        Message.user_id == user_id
    ).order_by(Message.timestamp.desc()).limit(500).all()

    return [
        {
            "id": m.id,
            "sender": m.sender,
            "content": m.content,
            "timestamp": m.timestamp
        }
        for m in reversed(messages)
    ]

@app.get("/admin/chat-list")
def get_chat_list(db: Session = Depends(get_db)):
    user_ids = db.query(Message.user_id).filter(Message.sender.in_(["user", "bot", "admin"])).distinct().all()
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

# ✅ 데이터베이스 테이블 생성 (최초 1회 자동 생성용)
Base.metadata.create_all(bind=engine)

import os

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port)

