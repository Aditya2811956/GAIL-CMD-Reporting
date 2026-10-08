from typing import Optional, List
from pydantic import BaseModel

class MicrosoftLoginRequest(BaseModel):
    microsoftId: str
    email: str
    name: str
    tenantId: Optional[str] = None
    mobilePhone: Optional[str] = None
    officeLocation: Optional[str] = None
    jobTitle: Optional[str] = None
    givenName: Optional[str] = None
    surname: Optional[str] = None
    preferredLanguage: Optional[str] = None
    businessPhones: Optional[List[str]] = None

class UserOut(BaseModel):
    id: str
    microsoftId: str
    name: str
    email: str
    tenantId: Optional[str] = None
    mobilePhone: Optional[str] = None
    officeLocation: Optional[str] = None
    jobTitle: Optional[str] = None
    givenName: Optional[str] = None
    surname: Optional[str] = None
    preferredLanguage: Optional[str] = None
    businessPhones: Optional[List[str]] = None

class TokenPair(BaseModel):
    access_token: str
    refresh_token: str
    expires_in: int

class AuthResponse(BaseModel):
    success: bool
    message: str
    is_new_user: bool
    user: UserOut
    tokens: TokenPair

class RefreshRequest(BaseModel):
    refresh_token: str
