# database.py
from motor.motor_asyncio import AsyncIOMotorClient
from bson.objectid import ObjectId
from dotenv import load_dotenv
import os
load_dotenv()

mongo_uri = os.getenv("MONGODB_URI")
client = AsyncIOMotorClient(mongo_uri)
db = client["company_employee_db"]

def get_company_collection():
    return db["companies"]

def get_employee_collection():
    return db["employees"]