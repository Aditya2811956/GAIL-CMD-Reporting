from typing import Optional
from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer

from core.security import decode_token

_bearer = HTTPBearer(auto_error=False)

async def get_current_user(
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
) -> dict:
    """
    Authenticate user via Bearer token or provide mock admin user for dev bypass.
    """
    if credentials and credentials.credentials:
        try:
            payload = decode_token(credentials.credentials)
            if payload.get("type") == "access":
                return payload
        except Exception:
            pass

    # Dev bypass default user
    return {
        "sub": "admin_user",
        "username": "admin",
        "email": "admin@gail.co.in",
        "role": "admin",
        "type": "access",
        "is_authenticated": True,
    }
