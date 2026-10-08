from datetime import datetime, timezone
from typing import Optional, Tuple
from bson import ObjectId
from pymongo.collection import Collection

def upsert_user(
    users: Collection,
    user_data: dict,
) -> Tuple[dict, bool]:
    now = datetime.now(timezone.utc)
    microsoft_id = user_data["microsoftId"]
    existing = users.find_one({"microsoftId": microsoft_id})
    if existing:
        update_data = {
            **user_data,
            "lastLogin": now,
        }
        users.update_one(
            {"_id": existing["_id"]},
            {"$set": update_data},
        )
        existing.update(update_data)
        return existing, False
    doc = {
        **user_data,
        "refreshToken": None,
        "createdAt": now,
        "lastLogin": now,
    }
    result = users.insert_one(doc)
    doc["_id"] = result.inserted_id
    return doc, True

def get_user_by_id(users: Collection, user_id: str) -> Optional[dict]:
    try:
        return users.find_one({"_id": ObjectId(user_id)})
    except Exception:
        return None

def save_refresh_token(users: Collection, user_id: str, token: str) -> None:
    users.update_one({"_id": ObjectId(user_id)}, {"$set": {"refreshToken": token}})

def validate_stored_refresh_token(user: dict, token: str) -> bool:
    return user.get("refreshToken") == token

def revoke_refresh_token(users: Collection, user_id: str) -> None:
    users.update_one({"_id": ObjectId(user_id)}, {"$set": {"refreshToken": None}})
