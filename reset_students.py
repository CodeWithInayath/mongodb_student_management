from pymongo import MongoClient

client = MongoClient("mongodb://localhost:27017/")
db = client["college"]
collection = db["students"]

result = collection.delete_many({})

print("Deleted records:", result.deleted_count)