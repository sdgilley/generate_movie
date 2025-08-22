from typing import Optional
from sqlmodel import SQLModel, Field

class User(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    username: str = Field(index=True)
    hashed_password: str
    foundry_endpoint_encrypted: Optional[str] = None
    foundry_key_encrypted: Optional[str] = None
    refresh_token_encrypted: Optional[str] = None
    msal_token_cache_encrypted: Optional[str] = None
