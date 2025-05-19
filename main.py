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

# 시트 이름 명시
SHEET_NAME = "통합관리"

# OneDrive Graph URL (Excel 파일 다운로드)
EXCEL_FILE_URL = "https://graph.microsoft.com/v1.0/sites/satmoulab.sharepoint.com:/sites/rental_data:/drive/items/AB8C3088-F264-49AD-8F8A-3D5723124A39/content"

# 엑세스 토큰 요청
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
def get_user_info(phone: str = Query(..., description="전화번호 예: 010-1234-5678")):
    try:
        token = get_access_token()

        headers = {"Authorization": f"Bearer {token}"}
        file_response = requests.get(EXCEL_FILE_URL, headers=headers)
        file_response.raise_for_status()

        # Excel 파일 읽기
        df = pd.read_excel(io.BytesIO(file_response.content), sheet_name=SHEET_NAME)
        df = df.fillna("")

        # J열(9), K열(10)에 전화번호 일치 여부 확인
        match_df = df[(df.iloc[:, 9] == phone) | (df.iloc[:, 10] == phone)]

        if len(match_df) == 0:
            return {"status": "not_found", "message": "해당 번호로 등록된 정보가 없습니다."}

        # 동일 번호 여러 개인 경우 Q열(16)이 비어 있는 행만 필터링
        if len(match_df) > 1:
            match_df = match_df[match_df.iloc[:, 16] == ""]

        if match_df.empty:
            return {"status": "not_found", "message": "일치하는 정보가 없습니다 (조건 불충족)."}

        row = match_df.iloc[0]
        return {
            "status": "ok",
            "name": str(row.iloc[8]).strip(),                   # I열: 수취인명
            "start_date": str(row.iloc[13]).split("T")[0],     # N열: 시작일
            "end_date": str(row.iloc[14]).split("T")[0],       # O열: 종료일
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))  # Render 환경에서 포트 자동 지정
    uvicorn.run("main:app", host="0.0.0.0", port=port)

# rebuild trigger dummy line


