from fastapi.middleware.cors import CORSMiddleware

def apply_cors(app):
    origins = [
        "http://localhost:5500",
        "http://127.0.0.1:5500",
        "http://127.0.0.1:8080",  # ✅ live-server용 추가
        "https://taupe-fox-b0ad0f.netlify.app",
        "https://idyllic-lamington-08426a.netlify.app",
        "https://cheerful-dolphin-1519ed.netlify.app",
        "https://lighthearted-cucurucho-c8d1c4.netlify.app",
        "https://celebrated-donut-5579ea.netlify.app"  # ✅ 현재 운영 주소
    ]

    app.add_middleware(
        CORSMiddleware,
        allow_origins=origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"]
    )




