from pymongo import MongoClient

# Connect to MongoDB
client = MongoClient("mongodb://localhost:27017/")

# Create database and collection
db = client["college"]
collection = db["students"]

# Add student
def add_student():
    name = input("Enter name: ")
    age = int(input("Enter age: "))
    course = input("Enter course: ")

    student = {
        "name": name,
        "age": age,
        "course": course
    }

    collection.insert_one(student)
    print("Student added successfully!")

# View students
def view_students():
    students = collection.find()
    print("\n--- Student List ---")
    for s in students:
        print(s)

# Update student
def update_student():
    name = input("Enter student name to update: ")
    new_age = int(input("Enter new age: "))

    collection.update_one(
        {"name": name},
        {"$set": {"age": new_age}}
    )

    print("Student updated!")

# Delete student
def delete_student():
    name = input("Enter student name to delete: ")

    collection.delete_one({"name": name})
    print("Student deleted!")

# Menu
while True:
    print("\n1. Add Student")
    print("2. View Students")
    print("3. Update Student")
    print("4. Delete Student")
    print("5. Exit")

    choice = int(input("Enter choice: "))

    if choice == 1:
        add_student()
    elif choice == 2:
        view_students()
    elif choice == 3:
        update_student()
    elif choice == 4:
        delete_student()
    elif choice == 5:
        break
    else:
        print("Invalid choice!")
