from fastapi import FastAPI, Query, Request, Depends, Response
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
import requests
import pandas as pd
import os
from datetime import datetime
from models import Base, Message, MessageCreate
from database import engine, SessionLocal
from typing import List
from pydantic import BaseModel

app = FastAPI()

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 환경변수
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")
PORT = int(os.environ.get("PORT", 10000))

# 고정값
SHAREPOINT_SITE_ID = "satmoulab.sharepoint.com,102fbb5d-7970-47e4-8686-f6d7fac0375f,cac8f27f-7023-4427-a96f-bd777b42c781"
EXCEL_ITEM_ID = "01BRDK2MMIGCGKWZHSVVEY7CR5K4RRESRZ"
SHEET_NAME = "통합관리"
RANGE_ADDRESS = "A1:Q30000"
DAILY_LATE_FEE = 1000  # 1일 연체료

class PhoneRequest(BaseModel):
    phone: str

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
    print("\ud83d\udccd Excel 요청 응답:", response.status_code, response.text)
    data = response.json()

    values = data.get("values", [])
    if not values or len(values) < 2:
        raise ValueError("\u274c 데이터 없음: 엑셀에서 값을 가져오지 못했습니다.")

    header = [str(h).strip() for h in values[0]]
    header_map = {h: i for i, h in enumerate(header)}
    rows = values[1:]

    print("\ud83d\udccc 헤더 확인:", header)

    try:
        contact1_idx = header_map["연락처1"]
        contact2_idx = header_map["연락처2"]
        name_idx = header_map["수취인명"]
        start_idx = header_map["시작일"]
        end_idx = header_map["종료일"]
        model_idx = header_map["제품명"]
        return_idx = header_map["반납완료일"]
    except KeyError as e:
        return {"error": f"필수 열 불러오기 실패: {e}"}

    phone = normalize_phone(phone)
    today = datetime.today().strftime("%Y-%m-%d")

    for row in reversed(rows):
        if len(row) < len(header):
            continue

        contact1 = normalize_phone(row[contact1_idx]) if contact1_idx < len(row) else ""
        contact2 = normalize_phone(row[contact2_idx]) if contact2_idx < len(row) else ""
        is_returned = row[return_idx] if return_idx < len(row) else None

        if phone == contact1 or phone == contact2:
            name = row[name_idx] if name_idx < len(row) else ""
            start = row[start_idx] if start_idx < len(row) else ""
            end = row[end_idx] if end_idx < len(row) else ""
            model = row[model_idx] if model_idx < len(row) else ""

            start_date = parse_excel_date(start)
            end_date = parse_excel_date(end)
            is_late = False
            late_days = 0
            late_fee = 0

            if not is_returned and end_date < today:
                is_late = True
                late_days = (datetime.strptime(today, "%Y-%m-%d") - datetime.strptime(end_date, "%Y-%m-%d")).days
                late_fee = late_days * DAILY_LATE_FEE

            return {
                "대여자명": name,
                "대여시작일": start_date,
                "대여종료일": end_date,
                "제품명": model,
                "연체여부": "Y" if is_late else "N",
                "연체일수": late_days,
                "연체료": late_fee
            }

    return {
        "대여자명": None,
        "대여시작일": None,
        "대여종료일": None,
        "제품명": None,
        "연체여부": None,
        "연체일수": None,
        "연체료": None
    }

@app.get("/")
def root():
    result = get_site_id_from_graph()
    return {"message": "FastAPI Excel 연결 OK", "site_id": result}

@app.post("/get-user-info")
async def get_user_info(req: PhoneRequest):
    try:
        phone = req.phone
        if not phone:
            return {"error": "전화번호가 누락되었습니다."}

        result = get_excel_data(phone)
        return jsonable_encoder(result)  # ← 이 줄로 변경
    except Exception as e:
        print("\u274c get-user-info 오류 발생:", str(e))
        return {"error": f"내부 오류: {str(e)}"}

# 입금 webhook
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

# DB 세션
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
    ).order_by(Message.timestamp.asc()).limit(500).all()
    return [
        {
            "id": m.id,
            "sender": m.sender,
            "content": m.content,
            "timestamp": m.timestamp
        }
        for m in messages
    ]

@app.get("/admin/chat-list")
def get_chat_list(db: Session = Depends(get_db)):
    user_ids = db.query(Message.user_id).filter(
        Message.sender.in_(["user", "bot", "admin"])
    ).distinct().all()
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

Base.metadata.create_all(bind=engine)

# site-id 확인용 함수
def get_site_id_from_graph():
    token = get_access_token()
    url = "https://graph.microsoft.com/v1.0/sites/satmoulab.sharepoint.com:/sites/rental_data"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    print("\ud83d\udccd site-id 결과:", response.status_code, response.text)
    return response.json()

@app.head("/", include_in_schema=False)
async def root_head():
    return Response(status_code=200)

