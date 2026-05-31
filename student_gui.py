from tkinter import *
from pymongo import MongoClient
from tkinter import messagebox

# Connect MongoDB
client = MongoClient("mongodb://localhost:27017/")
db = client["college"]
collection = db["students"]

# Create window
root = Tk()
root.title("Student Management System")
root.geometry("400x400")

# Labels
Label(root, text="Name").pack()
name_entry = Entry(root)
name_entry.pack()

Label(root, text="Age").pack()
age_entry = Entry(root)
age_entry.pack()

Label(root, text="Course").pack()
course_entry = Entry(root)
course_entry.pack()


# Add student
def add_student():
    name = name_entry.get()
    age = age_entry.get()
    course = course_entry.get()

    if name == "" or age == "" or course == "":
        messagebox.showwarning("Warning", "Fill all fields")
        return

    collection.insert_one({
        "name": name,
        "age": int(age),
        "course": course
    })

    messagebox.showinfo("Success", "Student Added")


# View students
def view_students():
    output.delete("1.0", END)

    students = collection.find()

    for s in students:
        output.insert(END,
                      f"Name: {s['name']} | Age: {s['age']} | Course: {s['course']}\n")


# Update student
def update_student():
    name = name_entry.get()
    age = age_entry.get()

    collection.update_one(
        {"name": name},
        {"$set": {"age": int(age)}}
    )

    messagebox.showinfo("Success", "Student Updated")


# Delete student
def delete_student():
    name = name_entry.get()

    collection.delete_one({"name": name})

    messagebox.showinfo("Success", "Student Deleted")


# Buttons
Button(root, text="Add Student", command=add_student).pack(pady=5)

Button(root, text="View Students", command=view_students).pack(pady=5)

Button(root, text="Update Student", command=update_student).pack(pady=5)

Button(root, text="Delete Student", command=delete_student).pack(pady=5)


# Output box
output = Text(root, height=10, width=45)
output.pack()


root.mainloop()