from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

DATABASE_URL = "postgresql://chatbot_db_pgh2_user:9U0tIEIBH5ErlXD2bToqIfWlznNRfupu@dpg-d0s22ie3jp1c73e8eos0-a.oregon-postgres.render.com/chatbot_db_pgh2"

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
