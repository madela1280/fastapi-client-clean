from fastapi import FastAPI, Query, Request
from fastapi.middleware.cors import CORSMiddleware
import requests
import pandas as pd
import os
from datetime import datetime, timedelta
import threading
import time

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

# 전역 캐시
cached_data = {
    "timestamp": None,
    "rows": [],
    "header": []
}
CACHE_TTL = 60  # 60초 주기

# 보조 함수들
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
    res = requests.post(url, headers=headers, data=data)
    return res.json().get("access_token")

def refresh_excel_cache():
    global cached_data
    print("🔄 Excel 캐시 갱신 시도 중...")
    token = get_access_token()
    if not token:
        print("❌ 토큰 발급 실패")
        return
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{EXCEL_ITEM_ID}/workbook/worksheets('{SHEET_NAME}')/range(address='{RANGE_ADDRESS}')"
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(url, headers=headers)
    json_data = res.json()
    if "values" not in json_data:
        print("❌ Excel 범위 오류:", json_data)
        return
    values = json_data["values"]
    cached_data["timestamp"] = datetime.utcnow()
    cached_data["header"] = values[0]
    cached_data["rows"] = values[1:]
    print(f"✅ 캐시 갱신 완료. 총 {len(cached_data['rows'])}행")

def cache_worker():
    while True:
        refresh_excel_cache()
        time.sleep(CACHE_TTL)

# 캐시 시작
threading.Thread(target=cache_worker, daemon=True).start()

# API 엔드포인트
@app.get("/")
def root():
    return {"message": "FastAPI Excel 연결 OK (캐싱 포함)"}

@app.get("/get-user-info")
def get_user_info(phone: str = Query(...)):
    phone = normalize_phone(phone)
    header = cached_data["header"]
    rows = cached_data["rows"]
    
    try:
        contact1_idx = header.index("연락처1")
        contact2_idx = header.index("연락처2")
        name_idx = header.index("수취인명")
        start_idx = header.index("시작일")
        end_idx = header.index("종료일")
        model_idx = header.index("제품명")
        return_idx = header.index("반납완료일") if "반납완료일" in header else None
    except ValueError as e:
        return {"error": f"열 이름 오류: {e}"}

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


