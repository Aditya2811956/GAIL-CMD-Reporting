import logging
from fastapi import APIRouter, Depends, HTTPException, status
from pymongo.errors import DuplicateKeyError

from configs.db import users
from core.security import (
    create_access_token,
    create_refresh_token,
    decode_token,
    ACCESS_TOKEN_EXPIRE_SECONDS,
)
from auth.schemas import (
    AuthResponse,
    MicrosoftLoginRequest,
    RefreshRequest,
    TokenPair,
    UserOut,
)
from auth.service import (
    get_user_by_id,
    revoke_refresh_token,
    save_refresh_token,
    upsert_user,
    validate_stored_refresh_token,
)
from auth.dependencies import get_current_user

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/auth", tags=["auth"])


def _build_user_out(user: dict) -> UserOut:
    return UserOut(
    id=str(user["_id"]),
    microsoftId=user["microsoftId"],
    name=user["name"],
    email=user["email"],
    tenantId=user.get("tenantId"),

    mobilePhone=user.get("mobilePhone"),
    officeLocation=user.get("officeLocation"),
    jobTitle=user.get("jobTitle"),
    givenName=user.get("givenName"),
    surname=user.get("surname"),
    preferredLanguage=user.get("preferredLanguage"),
    businessPhones=user.get("businessPhones"),
)


def _issue_tokens(user: dict) -> TokenPair:
    user_id = str(user["_id"])
    access_token = create_access_token(user_id, user["microsoftId"], user["email"])
    refresh_token = create_refresh_token(user_id)
    save_refresh_token(users, user_id, refresh_token)
    return TokenPair(
        access_token=access_token,
        refresh_token=refresh_token,
        expires_in=ACCESS_TOKEN_EXPIRE_SECONDS,
    )


#Routes

@router.post("/microsoft", response_model=AuthResponse)
async def microsoft_login(req: MicrosoftLoginRequest):
    """
    Called from the frontend immediately after MSAL resolves the account.
    Upserts the user in MongoDB and returns a fresh JWT pair.
    """
    try:
        user, is_new = upsert_user(
        users,
        req.model_dump(exclude_none=True),
        )
    except DuplicateKeyError:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="An account with this email already exists.",
        )
    except Exception:
        logger.exception("Microsoft login failed for %s", req.email)
        raise HTTPException(status_code=500, detail="Internal server error")

    tokens = _issue_tokens(user)

    return AuthResponse(
        success=True,
        message="Account created successfully" if is_new else "Welcome back",
        is_new_user=is_new,
        user=_build_user_out(user),
        tokens=tokens,
    )


@router.post("/refresh", response_model=TokenPair)
async def refresh_tokens(req: RefreshRequest):
    """
    Exchange a valid refresh token for a new access + refresh token pair.
    Call this on the frontend before the access token expires.
    """
    payload = decode_token(req.refresh_token)

    if payload.get("type") != "refresh":
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Expected a refresh token",
        )

    user = get_user_by_id(users, payload["sub"])
    if not user:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="User not found")

    if not validate_stored_refresh_token(user, req.refresh_token):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Refresh token has been revoked",
        )

    return _issue_tokens(user)


@router.post("/logout", status_code=status.HTTP_200_OK)
async def logout(current_user: dict = Depends(get_current_user)):
    """
    Revokes the stored refresh token.
    The access token stays valid until it naturally expires (stateless JWT).
    Frontend should discard both tokens from storage on receiving 200.
    """
    revoke_refresh_token(users, current_user["sub"])
    return {"success": True, "message": "Logged out successfully"}


@router.get("/me", response_model=UserOut)
async def get_me(current_user: dict = Depends(get_current_user)):
    """Returns the profile of the currently authenticated user."""
    user = get_user_by_id(users, current_user["sub"])
    if not user:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="User not found")
    return _build_user_out(user)
