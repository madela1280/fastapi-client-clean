from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
import requests
import io
import pandas as pd
import os

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

# OneDrive 파일 정보
EXCEL_FILE_URL = "https://graph.microsoft.com/v1.0/me/drive/items/02CEC702-0806-476E-AA5F-5C8BE1DAA19C/content"
SHEET_NAME = "통합관리"


def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, headers=headers, data=data)
    response.raise_for_status()
    return response.json().get("access_token")


@app.get("/get-user-info")
def get_user_info(phone: str = Query(..., description="전화번호, 예: 010-1234-5678")):
    try:
        # 토큰 발급
        token = get_access_token()

        # Excel 파일 다운로드
        headers = {"Authorization": f"Bearer {token}"}
        file_response = requests.get(EXCEL_FILE_URL, headers=headers)
        file_response.raise_for_status()

        # 엑셀 데이터 읽기
        df = pd.read_excel(io.BytesIO(file_response.content), sheet_name=SHEET_NAME)

        # 전화번호 양쪽 열에서 찾기 (J열=연락처1, K열=연락처2)
        df = df.fillna("")  # 결측치 방지
        match_df = df[(df.iloc[:, 9] == phone) | (df.iloc[:, 10] == phone)]  # J=9, K=10

        if len(match_df) == 0:
            return {"status": "not_found", "message": "해당 번호로 등록된 정보가 없습니다."}

        if len(match_df) > 1:
            match_df = match_df[match_df.iloc[:, 16] == ""]  # Q열(16번)이 빈 행만

        row = match_df.iloc[0]
        return {
            "status": "ok",
            "name": str(row.iloc[8]).strip(),           # I열: 수취인명
            "start_date": str(row.iloc[13]).split("T")[0],  # N열: 시작일
            "end_date": str(row.iloc[14]).split("T")[0],    # O열: 종료일
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}




