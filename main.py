from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
import requests
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

# SharePoint 및 Excel 정보
SHAREPOINT_SITE_ID = "satmoulab.sharepoint.com,102fbb5d-7970-47e4-8686-f6d7fac0375f,cac8f27f-7023-4427-a96f-bd777b42c781"
EXCEL_ITEM_ID = "01BRDK2MMIGCGKWZHSVVEY7CR5K4RRESRZ"
SHEET_NAME = "통합관리"
RANGE_ADDRESS = "H1:Q30000"  # 최대 3만 행까지 커버

# 전화번호 정규화
def normalize_phone(p):
    return str(p).replace("-", "").replace(" ", "").strip()

# 날짜 포맷 변환
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

# 고객 정보 조회
def get_excel_data(phone: str):
    token = get_access_token()
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{EXCEL_ITEM_ID}/workbook/worksheets('{SHEET_NAME}')/range(address='{RANGE_ADDRESS}')"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    data = response.json()

    values = data.get("values")
    if not values:
        print("❌ Excel 범위에서 값을 가져오지 못했습니다.")
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
        model_idx = header.index("제품명")  # 🔄 H열 기준으로 변경됨
        return_idx = header.index("반납완료일") if "반납완료일" in header else None
    except ValueError as e:
        print("❌ 열 이름이 일치하지 않음:", e)
        return None

    # 아래에서부터 검색하여 최신 행 선택
    for row in reversed(rows):
        contact1 = normalize_phone(row[contact1_idx]) if contact1_idx < len(row) else ""
        contact2 = normalize_phone(row[contact2_idx]) if contact2_idx < len(row) else ""
        is_returned = row[return_idx] if return_idx is not None and len(row) > return_idx else None

        if phone == contact1 or phone == contact2:
            if not is_returned:
                name = row[name_idx]
                start = row[start_idx]
                end = row[end_idx]
                model = row[model_idx] if model_idx < len(row) else ""
                return {
                    "대여자명": name,
                    "대여시작일": parse_excel_date(start),
                    "대여종료일": parse_excel_date(end),
                    "제품명": model
                }
    return None

# 루트 확인
@app.get("/")
def root():
    return {"message": "FastAPI Excel 연결 OK"}

# 고객 조회 API
@app.get("/get-user-info")
def get_user_info(phone: str = Query(..., description="전화번호('-' 없이) 입력")):
    result = get_excel_data(phone)
    if result:
        return result
    return {"message": "해당 전화번호로 등록된 정보가 없습니다."}