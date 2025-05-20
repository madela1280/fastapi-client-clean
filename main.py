from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
import requests
import io
import pandas as pd
import os
from datetime import datetime

app = FastAPI()

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 환경 변수에서 가져오기
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")

# 엑셀 경로 세팅
SHAREPOINT_SITE_ID = "your_site_id"
EXCEL_ITEM_ID = "your_excel_file_item_id"
SHEET_NAME = "통합관리"

# 엑세스 토큰 발급
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

# 엑셀 파일에서 정보 가져오기
def get_excel_data(phone: str):
    token = get_access_token()
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{EXCEL_ITEM_ID}/workbook/worksheets('{SHEET_NAME}')/usedRange"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    data = response.json()

    values = data.get("values", [])
    header = values[0]
    rows = values[1:]

    phone = phone.replace("-", "").strip()
    contact1_idx = header.index("연락처1")
    contact2_idx = header.index("연락처2")
    name_idx = header.index("수취인명")
    start_idx = header.index("대여시작일")
    end_idx = header.index("대여종료일")
    return_idx = header.index("반납일") if "반납일" in header else None

    for row in rows:
        contact1 = str(row[contact1_idx]).replace("-", "").strip()
        contact2 = str(row[contact2_idx]).replace("-", "").strip()
        is_returned = row[return_idx] if return_idx is not None and len(row) > return_idx else None

        if phone == contact1 or phone == contact2:
            if not is_returned:
                name = row[name_idx]
                start = row[start_idx]
                end = row[end_idx]
                # 날짜 포맷 정제
                start_date = parse_excel_date(start)
                end_date = parse_excel_date(end)
                return {
                    "name": name,
                    "start_date": start_date,
                    "end_date": end_date,
                }

    return None

# 날짜 포맷 처리 함수
def parse_excel_date(value):
    if isinstance(value, float) or isinstance(value, int):
        base_date = datetime(1899, 12, 30)
        return (base_date + pd.to_timedelta(value, unit="D")).strftime("%Y-%m-%d")
    if isinstance(value, str):
        try:
            return datetime.strptime(value[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        except:
            return value
    return str(value)

# ✅ 루트 경로 응답 추가
@app.get("/")
def root():
    return {"message": "FastAPI app is running on Render!"}

# 📞 /get-user-info?phone=01012345678
@app.get("/get-user-info")
def get_user_info(phone: str = Query(..., description="전화번호를 '-' 없이 입력")):
    result = get_excel_data(phone)
    if result:
        return result
    return {"message": "해당 전화번호로 등록된 정보가 없습니다."}




