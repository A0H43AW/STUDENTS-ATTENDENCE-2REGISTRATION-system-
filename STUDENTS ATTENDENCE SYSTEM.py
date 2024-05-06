import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import os

# Function to handle image upload
def upload_image():
    filename = filedialog.askopenfilename(title="Upload Image", filetypes=[("Image files", "*.jpg *.jpeg *.png")])
    if filename:
        load_image(filename)

# Function to load and display the image
def load_image(filename):
    global img_label, img_data
    img = Image.open(filename)
    img = img.resize((100, 100), Image.ANTIALIAS)
    img_data = ImageTk.PhotoImage(img)
    img_label.config(image=img_data)

# Function to handle the registration process
def register_attendance():
    # Retrieve student details from the input fields
    name = name_entry.get()
    reg_number = reg_entry.get()
    lecture_number = lecture_entry.get()

    # Check if any field is empty
    if not name or not reg_number or not lecture_number:
        messagebox.showerror("Error", "Please fill in all fields")
        return

    # Append student details to the Excel sheet
    data = {'Name': [name], 'Registration Number': [reg_number], 'Lecture Number': [lecture_number]}
    df = pd.DataFrame(data)
    if os.path.exists('attendance.xlsx'):
        df.to_excel('attendance.xlsx', index=False, header=False, mode='a')
    else:
        df.to_excel('attendance.xlsx', index=False)

    # Show success message
    messagebox.showinfo("Success", "Attendance registered successfully")

# Function to clear all fields
def clear_fields():
    name_entry.delete(0, tk.END)
    reg_entry.delete(0, tk.END)
    lecture_entry.delete(0, tk.END)
    img_label.config(image="")

# Create the main window
root = tk.Tk()
root.title("Student Attendance Registration")

# Labels
tk.Label(root, text="Name:").grid(row=0, column=0, padx=10, pady=5)
tk.Label(root, text="Registration Number:").grid(row=1, column=0, padx=10, pady=5)
tk.Label(root, text="Lecture Number:").grid(row=2, column=0, padx=10, pady=5)
tk.Label(root, text="Profile Picture:").grid(row=3, column=0, padx=10, pady=5)

# Entry fields
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1, padx=10, pady=5)
reg_entry = tk.Entry(root)
reg_entry.grid(row=1, column=1, padx=10, pady=5)
lecture_entry = tk.Entry(root)
lecture_entry.grid(row=2, column=1, padx=10, pady=5)

# Profile Picture Upload button
upload_button = tk.Button(root, text="Upload", command=upload_image)
upload_button.grid(row=3, column=1, padx=10, pady=5)

# Image label
img_label = tk.Label(root)
img_label.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

# Register button
register_button = tk.Button(root, text="Register", command=register_attendance)
register_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

# Clear button
clear_button = tk.Button(root, text="Clear", command=clear_fields)
clear_button.grid(row=6, column=0, columnspan=2, padx=10, pady=5)

# Start the GUI application
root.mainloop()
