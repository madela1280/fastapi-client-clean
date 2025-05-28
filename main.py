from fastapi import FastAPI, Query, Request
from fastapi.middleware.cors import CORSMiddleware
import httpx
import pandas as pd
from datetime import datetime
import os
from time import time

app = FastAPI()

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
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

# 캐시 구조 (60초 유지)
_excel_cache = {"data": None, "last_fetched": 0}
CACHE_DURATION = 60

# 전화번호 정규화
def normalize_phone(p):
    return str(p).replace("-", "").replace(" ", "").strip()

# 엑셀 날짜 파싱
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

# 액세스 토큰 캐시
access_token_cache = {"token": None, "expires_at": 0}

async def get_access_token():
    if access_token_cache["token"] and access_token_cache["expires_at"] > datetime.now().timestamp():
        return access_token_cache["token"]

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    async with httpx.AsyncClient() as client:
        res = await client.post(url, data=data, headers=headers)
        res.raise_for_status()
        token = res.json()["access_token"]
        access_token_cache["token"] = token
        access_token_cache["expires_at"] = datetime.now().timestamp() + 3400
        return token

# 핵심 로직: 캐시 포함된 엑셀 조회
async def get_excel_data(phone: str):
    now = time()
    if _excel_cache["data"] and now - _excel_cache["last_fetched"] < CACHE_DURATION:
        values = _excel_cache["data"]
    else:
        token = await get_access_token()
        url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{EXCEL_ITEM_ID}/workbook/worksheets('{SHEET_NAME}')/range(address='{RANGE_ADDRESS}')"
        headers = {"Authorization": f"Bearer {token}"}
        async with httpx.AsyncClient(timeout=20.0) as client:
            res = await client.get(url, headers=headers)
            res.raise_for_status()
            values = res.json().get("values")
        _excel_cache["data"] = values
        _excel_cache["last_fetched"] = now

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

# 루트 확인용
@app.get("/")
def root():
    return {"message": "FastAPI Excel 연결 OK"}

# 고객 조회 엔드포인트
@app.get("/get-user-info")
async def get_user_info(phone: str = Query(...)):
    return await get_excel_data(phone)

# ✅ 입금 문자 Webhook (누락되었던 부분 복구)
@app.post("/deposit-webhook")
async def handle_sms_webhook(request: Request):
    body = await request.body()
    content = body.decode("utf-8")
    today = datetime.now().strftime("%m/%d")

    if today not in content:
        return {"message": "오늘 날짜 문자 아님"}

    return {"message": "입금 문자 수신됨", "본문": content}






