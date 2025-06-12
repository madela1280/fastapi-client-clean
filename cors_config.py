from fastapi.middleware.cors import CORSMiddleware

def apply_cors(app):
    origins = [
        "http://localhost:5500",
        "http://127.0.0.1:5500",
        "https://taupe-fox-b0ad0f.netlify.app"  # ✅ 현재 Netlify URL
    ]

    app.add_middleware(
        CORSMiddleware,
        allow_origins=origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"]
    )


