from fastapi import FastAPI, Query, Request
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import os
from datetime import datetime

app = FastAPI()

# ✅ CORS 설정 (Netlify 주소 등 정확히 명시)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://genuine-treacle-599cab.netlify.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 📄 엑셀 경로 지정 (Render에 업로드된 경우 static 경로 등)
EXCEL_PATH = "local_data.xlsx"  # 실제 서버에 배포한 엑셀 파일 경로

# 🧠 메모리 캐시 초기화
cached_rows = []

# 📆 날짜 파싱 함수
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

# ☎️ 전화번호 정규화
def normalize_phone(p):
    return str(p).replace("-", "").replace(" ", "").strip()

# ✅ 서버 시작 시 엑셀을 메모리에 캐시
def load_excel():
    global cached_rows
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name="통합관리", dtype=str)
        df.fillna("", inplace=True)
        cached_rows = df.to_dict(orient="records")
        print(f"✅ {len(cached_rows)} rows loaded from Excel.")
    except Exception as e:
        print("❌ Excel 로딩 실패:", e)

@app.on_event("startup")
def startup_event():
    load_excel()

# 📦 GET 요청 처리
@app.get("/get-user-info")
def get_user_info(phone: str = Query(..., description="전화번호('-' 없이) 입력")):
    phone = normalize_phone(phone)

    for row in reversed(cached_rows):  # 최근 항목 우선 검색
        contact1 = normalize_phone(row.get("연락처1", ""))
        contact2 = normalize_phone(row.get("연락처2", ""))
        is_returned = row.get("반납완료일", "")

        if phone == contact1 or phone == contact2:
            if not is_returned:
                return {
                    "대여자명": row.get("수취인명", ""),
                    "대여시작일": parse_excel_date(row.get("시작일", "")),
                    "대여종료일": parse_excel_date(row.get("종료일", "")),
                    "제품명": row.get("제품명", ""),
                }

    return {
        "대여자명": None,
        "대여시작일": None,
        "대여종료일": None,
        "제품명": None
    }

# 🔍 상태 확인용
@app.get("/")
def root():
    return {"message": "FastAPI Excel 캐시 구조 OK"}







