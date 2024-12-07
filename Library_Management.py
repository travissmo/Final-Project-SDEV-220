#Library Management Project By Travis Smothermon for SDEV220 final project
#This program will use a database to store Book and Member info for a Library
#It will allow you to add book and member info and give other information about each

#Importing libraries. Tkinter for GUI and pandas for Excel/CSV functions
import tkinter as tk
from tkinter import messagebox
import pandas as pd 
import os

root = tk.Tk()
root.title("Library Management System")
root.geometry("900x400")

#Class setup
class Library:
    def __init__(self):
        self.books = []
        self.members = []

    def add_book(self, book):
        self.books.append(book)

    def add_member(self, member):
        self.members.append(member)

class Book:
    def __init__(self, title, author, isbn):
        self.title = title
        self.author = author
        self.isbn = isbn
        self.is_available = True

class Member:
    def __init__(self, name, member_id):
        self.name = name
        self.member_id = member_id

#Global variable for library
library = Library()

#Functions
#this function is what allows the addition of new books to excel sheet. This and the add_members function are identical 
def add_book():
    title = entry_title.get()
    author = entry_author.get()
    isbn = entry_isbn.get()

    #error checks for inputs in all fields, does not allow null data to be added
    if not title or not author or not isbn:
        messagebox.showerror("Input Error", "All fields must be filled out!")
        return

    new_book = Book(title, author, isbn)
    library.add_book(new_book)

    #Prepare book data for writing to the Excel file
    book_data = {
        "Title": [new_book.title],
        "Author": [new_book.author],
        "ISBN": [new_book.isbn],
        "Availability": ["Available"]
    }
    df_new_book = pd.DataFrame(book_data)
    file_path = os.path.join("data", "books_and_members.xlsx")
#had to get assistance with this part, still not exactly sure how it works but I believe
#this adds the excel functionality using a pandas class that allows writing to the excel sheet
    try:
        # Check if the file exists
        if os.path.exists(file_path):
            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                if 'Books' in writer.sheets:
                    df_new_book.to_excel(writer, index=False, header=False, startrow=writer.sheets['Books'].max_row, sheet_name='Books')
                else:
                    # If the 'Books' sheet doesn't exist, create it and write data
                    df_new_book.to_excel(writer, index=False, sheet_name='Books')
        else:
            df_new_book.to_excel(file_path, index=False, sheet_name='Books')

        #clears fields after hitting add button
        entry_title.delete(0, tk.END)
        entry_author.delete(0, tk.END)
        entry_isbn.delete(0, tk.END)
        messagebox.showinfo("Success", f"Book '{new_book.title}' added successfully!")

    except Exception as e:
        messagebox.showerror("File Error", f"An error occurred while saving the book: {e}")
#this function loads the data from the excel sheet "Books" onto tkinter window, similar to load_members
def load_books_from_excel():
    file_path = os.path.join("data", "books_and_members.xlsx")
    #allows overite data 
    try:
        if os.path.exists(file_path):
            df_books = pd.read_excel(file_path, sheet_name='Books')
            listbox_books.delete(0, tk.END)
            listbox_books.insert(tk.END, "Title | Author | ISBN | Availability")

            for index, row in df_books.iterrows():
                book_info = f"{row['Title']} by {row['Author']} (ISBN: {row['ISBN']}) - {row['Availability']}"
                listbox_books.insert(tk.END, book_info)
        else: #error check for if the book exists or not 
            messagebox.showwarning("Warning", "No book data found! Please add books first.")
    #error check
    except Exception as e:
        messagebox.showerror("File Error", f"An error occurred while loading books: {e}")
#this function is for adding members to "Members" sheet in excel, similar logic to "Books" code.
def add_member():
    name = entry_name.get()
    member_id = entry_member_id.get()

    #error checks to make sure that all fields are filled out, will not allow empty fields
    if not name or not member_id:
        messagebox.showerror("Input Error", "All fields must be filled out!")
        return

    new_member = Member(name, member_id)
    library.add_member(new_member)

    member_data = {"Name": [new_member.name], "Member ID": [new_member.member_id]}
    df_new_member = pd.DataFrame(member_data)
    file_path = os.path.join("data", "books_and_members.xlsx")

    try:
        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            if 'Members' in writer.sheets:
                df_new_member.to_excel(writer, index=False, header=False, startrow=writer.sheets['Members'].max_row, sheet_name='Members')
            else:
                df_new_member.to_excel(writer, index=False, sheet_name='Members')

        #Clear the text boxes after hitting the add button
        entry_name.delete(0, tk.END)
        entry_member_id.delete(0, tk.END)
        messagebox.showinfo("Success", f"Member '{new_member.name}' added successfully!")

    except Exception as e:
        messagebox.showerror("File Error", f"An error occurred while saving the member: {e}")
#loads the members info from excel in the tkinter window, similar to books.
def load_members():
    file_path = os.path.join("data", "books_and_members.xlsx")

    try:
        if os.path.exists(file_path):
            df_members = pd.read_excel(file_path, sheet_name="Members")
            listbox_members.delete(0, tk.END)
            listbox_members.insert(tk.END, "Name | Member ID")

            for _, row in df_members.iterrows():
                listbox_members.insert(tk.END, f"{row['Name']} (ID: {row['Member ID']})")
        else:
            messagebox.showwarning("Warning", "No member data found! Please add members first.")
    
    except Exception as e:
        messagebox.showerror("File Error", f"An error occurred while loading members: {e}")

#GUI structure
frame_books = tk.Frame(root, bd=2, relief="groove", padx=10, pady=10)
frame_books.grid(row=0, column=0, padx=10, pady=10)

tk.Label(frame_books, text="Books").grid(row=0, column=0, columnspan=2)

tk.Label(frame_books, text="Book Title:").grid(row=1, column=0, sticky="e")
entry_title = tk.Entry(frame_books)
entry_title.grid(row=1, column=1)

tk.Label(frame_books, text="Author:").grid(row=2, column=0, sticky="e")
entry_author = tk.Entry(frame_books)
entry_author.grid(row=2, column=1)

tk.Label(frame_books, text="ISBN:").grid(row=3, column=0, sticky="e")
entry_isbn = tk.Entry(frame_books)
entry_isbn.grid(row=3, column=1)

button_add_book = tk.Button(frame_books, text="Add Book", command=add_book)
button_add_book.grid(row=4, column=0, columnspan=2, pady=10)

listbox_books = tk.Listbox(frame_books, width=50, height=10)
listbox_books.grid(row=5, column=0, columnspan=2)

button_load_books = tk.Button(frame_books, text="Load Books", command=load_books_from_excel)
button_load_books.grid(row=6, column=0, columnspan=2)

#Members Section
frame_members = tk.Frame(root, bd=2, relief="groove", padx=10, pady=10)
frame_members.grid(row=0, column=1, padx=10, pady=10)

tk.Label(frame_members, text="Members").grid(row=0, column=0, columnspan=2)

tk.Label(frame_members, text="Name:").grid(row=1, column=0, sticky="e")
entry_name = tk.Entry(frame_members)
entry_name.grid(row=1, column=1)

tk.Label(frame_members, text="Member ID:").grid(row=2, column=0, sticky="e")
entry_member_id = tk.Entry(frame_members)
entry_member_id.grid(row=2, column=1)

button_add_member = tk.Button(frame_members, text="Add Member", command=add_member)
button_add_member.grid(row=3, column=0, columnspan=2, pady=10)

listbox_members = tk.Listbox(frame_members, width=50, height=10)
listbox_members.grid(row=4, column=0, columnspan=2)

button_load_members = tk.Button(frame_members, text="Load Members", command=load_members)
button_load_members.grid(row=5, column=0, columnspan=2)

root.mainloop()


