import os
from datetime import datetime, timedelta, timezone
from typing import Any, Dict
from jose import JWTError, jwt
from fastapi import HTTPException, status
SECRET_KEY = os.getenv("JWT_SECRET_KEY", "change-this-in-production")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_SECONDS  = 60 * 60
REFRESH_TOKEN_EXPIRE_SECONDS = 60 * 60 * 24 * 7

def create_access_token(user_id: str, microsoft_id: str, email: str) -> str:
    expire = datetime.now(timezone.utc) + timedelta(seconds=ACCESS_TOKEN_EXPIRE_SECONDS)
    return jwt.encode(
        {"sub": user_id, "microsoftId": microsoft_id, "email": email, "type": "access", "exp": expire},
        SECRET_KEY,
        algorithm=ALGORITHM,
    )
def create_refresh_token(user_id: str) -> str:
    expire = datetime.now(timezone.utc) + timedelta(seconds=REFRESH_TOKEN_EXPIRE_SECONDS)
    return jwt.encode(
        {"sub": user_id, "type": "refresh", "exp": expire},
        SECRET_KEY,
        algorithm=ALGORITHM,
    )
def decode_token(token: str) -> Dict[str, Any]:
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except JWTError:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or expired token",
            headers={"WWW-Authenticate": "Bearer"},
        )
