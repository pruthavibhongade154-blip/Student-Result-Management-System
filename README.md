# Student-Result-Management-System
This project is a Student Result Management System developed using Python. It allows a user to enter student details, calculate results, save records in an Excel file, and search for a student’s result using their roll number. The program provides a graphical user interface, so the user can enter information through text boxes and perform actions 




import tkinter as tk
from tkinter import messagebox
import openpyxl
import os

file = "student_results.xlsx"

if not os.path.exists(file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Name", "Roll No", "Class", "Sub1(maths)", "Sub2(english)",
    "Sub3(science)", "Sub4(hindi)", "Sub5(socialS)", "Total", "Percentage", "Result"])
    wb.save(file)

def save():
    try:
        marks = [int(s1.get()), int(s2.get()), int(s3.get()), int(s4.get()),
        int(s5.get())]

        total = sum(marks)
        per = total / 5
        result = "Pass" if per >= 40 else "Fail"

        wb = openpyxl.load_workbook(file)
        ws = wb.active

        ws.append([
            name.get(),
            roll.get(),
            cls.get(),
            *marks,
            total,
            per,
            result
        ])

        wb.save(file)

        messagebox.showinfo(
            "Result",
            f"Total: {total}\nPercentage: {per:.2f}%\nResult: {result}"
        )

    except:
        messagebox.showerror("Error", "Enter valid data")

def search():
    r = search_roll.get()

    wb = openpyxl.load_workbook(file)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[1]) == r:
            output.config(
                text=f"Name: {row[0]}\n"
                     f"Total: {row[8]}\n"
                     f"Percentage: {row[9]:.2f}%\n"
                     f"Result: {row[10]}"
            )
            return

    output.config(text="Student not found")


root = tk.Tk()
root.title("Student Result")
root.geometry("400x550")

tk.Label(root, text="Student Result Management",
         font=("Arial", 16, "bold")).pack(pady=10)

name = tk.Entry(root)
roll = tk.Entry(root)
cls = tk.Entry(root)
s1 = tk.Entry(root)
s2 = tk.Entry(root)
s3 = tk.Entry(root)
s4 = tk.Entry(root)
s5 = tk.Entry(root)

fields = [
    ("Name", name),
    ("Roll No", roll),
    ("Class", cls),
    ("Subject 1", s1),
    ("Subject 2", s2),
    ("Subject 3", s3),
    ("Subject 4", s4),
    ("Subject 5", s5)
]

for text, entry in fields:
    tk.Label(root, text=text).pack()
    entry.pack()

tk.Button(root, text="Save", command=save).pack(pady=10)

tk.Label(root, text="Enter Roll No to Search").pack()

search_roll = tk.Entry(root)
search_roll.pack()

tk.Button(root, text="Get Result", command=search).pack(pady=10)

output = tk.Label(root, text="", font=("Arial", 11))
output.pack(pady=10)

root.mainloop()
