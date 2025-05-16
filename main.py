from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import aiohttp
import asyncio
import re
from datetime import datetime

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 인증정보 설정
TENANT_ID = "8ff73382-61a3-420a-bc35-1f1969cf48db"
CLIENT_ID = "d2566ba2-91b2-42ca-a829-c39da8dfba3d"
CLIENT_SECRET = "<YOUR_CLIENT_SECRET>"  # 비밀값은 여기 직접 입력
EXCEL_FILE_ID = "02CEC702-0806-476E-AA5F-5C8BE1DAA19C"  # 파일 ID
SHEET_NAME = "통합관리"

# Microsoft Graph API 인증 토큰 발급
async def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(url, headers=headers, data=data) as resp:
            res = await resp.json()
            return res.get("access_token")

# 엑셀 데이터 조회
async def get_excel_data(token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{EXCEL_FILE_ID}/workbook/worksheets('{SHEET_NAME}')/usedRange"
    headers = {"Authorization": f"Bearer {token}"}
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            return await resp.json()

@app.post("/get-user-info")
async def get_user_info(request: Request):
    body = await request.json()
    phone_input = body.get("queryResult", {}).get("parameters", {}).get("phone-number", "")

    if isinstance(phone_input, list) and phone_input:
        phone_input = phone_input[0]

    digits = re.sub(r'[^0-9]', '', phone_input)
    if len(digits) == 11:
        formatted_input = f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    else:
        formatted_input = phone_input

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

                start_str = start if isinstance(start, str) else str(start)
                end_str = end if isinstance(end, str) else str(end)

                result = f"\ud83d\udce6 \ub300\uc5ec\uc790\uba85: {name}\n\ud83d\udcc5 \ub300\uc5ec\uc2dc\uc791\uc77c: {start_str}\n\u23f3 \ub300\uc5ec\uc885\ub8cc\uc77c: {end_str}"
                break

        if not result:
            result = "\uace0\uac1d \uc815\ubcf4\ub97c \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4.\\n\ub300\uc5ec \uc2dc \ub4f1\ub85d\ud55c \uc815\ud655\ud55c \uc804\ud654\ubc88\ud638\ub97c \uc785\ub825\ud574 \uc8fc\uc138\uc694."

        return JSONResponse(content={"fulfillmentText": result})

    except Exception as e:
        print("\u274c \uc624\ub958:", str(e))
        return JSONResponse(content={"fulfillmentText": "\uc2dc\uc2a4\ud15c \uc624\ub958\uac00 \ubc1c\uc0dd\ud588\uc2b5\ub2c8\ub2e4. \uc7a0\uc2dc \ud6c4 \ub2e4\uc2dc \uc2dc\ub3c4\ud574 \uc8fc\uc138\uc694."})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)




