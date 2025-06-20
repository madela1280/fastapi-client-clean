from sqlalchemy import Column, Integer, String, Text, DateTime, Boolean
from sqlalchemy.sql import func
from sqlalchemy.ext.declarative import declarative_base
from pydantic import BaseModel
from typing import Optional
from datetime import datetime

Base = declarative_base()

class Message(Base):
    __tablename__ = "messages"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(String)
    sender = Column(String)
    content = Column(Text)
    timestamp = Column(DateTime(timezone=False), server_default=func.now())
    read = Column(Boolean, default=False)

# ✅ 아래는 FastAPI 요청 처리용 Pydantic 모델
class MessageCreate(BaseModel):
    user_id: str
    role: str
    message: str
    timestamp: datetime
    read: Optional[bool] = False



