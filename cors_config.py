# cors_config.py
from fastapi.middleware.cors import CORSMiddleware

def apply_cors(app):
    origins = [
        "http://localhost:5500",
        "http://127.0.0.1:5500",
        "https://idyllic-lamington-08426a.netlify.app",
        "https://taupe-fox-b0ad0f.netlify.app"  # ✅ 새 Netlify 도메인 추가
    ]

    app.add_middleware(
        CORSMiddleware,
        allow_origins=origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"]
    )


