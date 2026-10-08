from pymongo import MongoClient
from pymongo.errors import PyMongoError
from dotenv import load_dotenv
import os
import sys

load_dotenv()

print(os.getenv("MONGO_URI"))

try:
    mongo_uri = os.getenv("MONGO_URI")
    if mongo_uri:
        client = MongoClient(
            mongo_uri,
            serverSelectionTimeoutMS=2000,
        )
        client.admin.command("ping")
        db = client["parliament"]
        users = db["users"]
        print("[OK] MongoDB Connected")
    else:
        client = None
        db = None
        users = None
        print("[INFO] MONGO_URI not set, auth database offline")

except Exception as e:
    client = None
    db = None
    users = None
    print(f"[WARN] MongoDB not available: {e}")
