<<<<<<< HEAD
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import aiohttp
import asyncio
import os
import re
from datetime import datetime
from pydantic import BaseModel
from typing import Optional, Union

app = FastAPI()

=======
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
import requests
import pandas as pd
import os
from datetime import datetime

app = FastAPI()

# CORS 설정
>>>>>>> 5a9be19958919d1b7ce45139fb02a6d4a0a0fe95
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

<<<<<<< HEAD
# 인증정보 환경변수
TENANT_ID = "8ff73382-61a3-420a-bc35-1f1969cf48db"
CLIENT_ID = "d2566ba2-91b2-42ca-a829-c39da8dfba3d"
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
EXCEL_FILE_ID = "02CEC702-0806-476E-AA5F-5C8BE1DAA19C"
SHEET_NAME = "통합관리"

# 요청 body 스키마
class Parameters(BaseModel):
    phone_number: Union[str, list]

class QueryResult(BaseModel):
    parameters: Parameters

class UserRequest(BaseModel):
    queryResult: Optional[QueryResult]

# Access Token 발급
async def get_access_token():
=======
# 환경 변수에서 가져오기
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")

# SharePoint 및 Excel 정보
SHAREPOINT_SITE_ID = "satmoulab.sharepoint.com,102fbb5d-7970-47e4-8686-f6d7fac0375f,cac8f27f-7023-4427-a96f-bd777b42c781"
EXCEL_ITEM_ID = "01BRDK2MMIGCGKWZHSVVEY7CR5K4RRESRZ"
SHEET_NAME = "통합관리"
RANGE_ADDRESS = "I1:Q500"  # 수취인명(I) ~ 반납완료일(Q)

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
>>>>>>> 5a9be19958919d1b7ce45139fb02a6d4a0a0fe95
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
<<<<<<< HEAD
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(url, headers=headers, data=data) as resp:
            res = await resp.json()
            return res.get("access_token")

# Excel 데이터 읽기
async def get_excel_data(token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{EXCEL_FILE_ID}/workbook/worksheets('{SHEET_NAME}')/usedRange"
    headers = {"Authorization": f"Bearer {token}"}
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            return await resp.json()

# POST 엔드포인트
@app.post("/get-user-info")
async def get_user_info(req: UserRequest):
    phone_input = req.queryResult.parameters.phone_number if req.queryResult else ""
    if isinstance(phone_input, list):
        phone_input = phone_input[0]
    digits = re.sub(r'[^0-9]', '', phone_input)
    formatted_input = f"{digits[:3]}-{digits[3:7]}-{digits[7:]}" if len(digits) == 11 else phone_input

    try:
        token = await get_access_token()
        data = await get_excel_data(token)
        values = data.get("values", [])

        headers = values[0] if values else []
        rows = values[1:] if len(values) > 1 else []

        result = None
        for row in rows:
            tel1 = str(row[9]) if len(row) > 9 else ""
            tel2 = str(row[10]) if len(row) > 10 else ""
            returned = row[16] if len(row) > 16 else ""

            if formatted_input in [tel1, tel2] and not returned:
                name = row[8] if len(row) > 8 else ""
                start = row[13] if len(row) > 13 else ""
                end = row[14] if len(row) > 14 else ""

                result = f"📦 대여자명: {name}\n📅 대여시작일: {start}\n⏳ 대여종료일: {end}"
                break

        if not result:
            result = "고객 정보를 찾을 수 없습니다.\n대여 시 등록한 정확한 전화번호를 입력해 주세요."

        return JSONResponse(content={"fulfillmentText": result})

    except Exception as e:
        print("❌ 오류:", str(e))
        return JSONResponse(content={"fulfillmentText": "시스템 오류가 발생했습니다. 잠시 후 다시 시도해 주세요."})

# 로컬 테스트용 (Render에서는 무시됨)
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)





=======
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
        return_idx = header.index("반납완료일") if "반납완료일" in header else None
    except ValueError as e:
        print("❌ 열 이름이 일치하지 않음:", e)
        return None

    for row in rows:
        contact1 = normalize_phone(row[contact1_idx]) if contact1_idx < len(row) else ""
        contact2 = normalize_phone(row[contact2_idx]) if contact2_idx < len(row) else ""
        is_returned = row[return_idx] if return_idx is not None and len(row) > return_idx else None

        if phone == contact1 or phone == contact2:
            if not is_returned:
                name = row[name_idx]
                start = row[start_idx]
                end = row[end_idx]
                return {
                    "대여자명": name,
                    "대여시작일": parse_excel_date(start),
                    "대여종료일": parse_excel_date(end)
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
>>>>>>> 5a9be19958919d1b7ce45139fb02a6d4a0a0fe95
