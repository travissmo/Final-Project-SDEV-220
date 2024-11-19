#Library Management Project By Travis Smothermon for SDEV220 final project
#This program will use a database to store Book and Member info for a Library
#It will allow you to add/edit book and member info and give other information about each

#Importing libraries. Tkinter for GUI and pandas for Excel/CSV functions
import tkinter as tk
from tkinter import messagebox
import pandas as pd 
import os

root = tk.Tk()
root.title("Library Management System")
root.geometry("600x400")

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
def add_book():
    title = entry_title.get()
    author = entry_author.get()
    isbn = entry_isbn.get()

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

    #This checks if the file exists, was having difficulty running the program until putting this in
    if os.path.exists(file_path):
        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            if 'Books' in writer.sheets:
                df_new_book.to_excel(writer, index=False, header=False, startrow=writer.sheets['Books'].max_row, sheet_name='Books')
            else:
                # If the 'Books' sheet doesn't exist, create it and write data
                df_new_book.to_excel(writer, index=False, sheet_name='Books')
    else:
        df_new_book.to_excel(file_path, index=False, sheet_name='Books')

    print(f"Book '{new_book.title}' added and saved to Excel.")

    #clears the fields after hitting the button
    entry_title.delete(0, tk.END)
    entry_author.delete(0, tk.END)
    entry_isbn.delete(0, tk.END)

def load_books_from_excel():
    
    file_path = os.path.join("data", "books_and_members.xlsx")
    
    
    if os.path.exists(file_path):
        
        df_books = pd.read_excel(file_path, sheet_name='Books')
        
        
        listbox_books.delete(0, tk.END)
        
        #This puts the headers on the window when you load the book data
        listbox_books.insert(tk.END, "Title | Author | ISBN | Availability")
        
        
        for index, row in df_books.iterrows():
            book_info = f"{row['Title']} by {row['Author']} (ISBN: {row['ISBN']}) - {row['Availability']}"
            listbox_books.insert(tk.END, book_info)

#Creates the window that shows the data in the excel file
def create_gui():
    global listbox_books
    

    
    #Displays the books inside the main Tkinter window
    listbox_books = tk.Listbox(root, width=50, height=10)
    listbox_books.grid(row=4, column=0, columnspan=2, pady=20)  # Adjust grid to fit in the layout

    load_button = tk.Button(root, text="Load Books", command=load_books_from_excel)
    load_button.grid(row=5, column=0, columnspan=2)  # Place it below the listbox

    #Start the main loop
    root.mainloop()

#GUI structure
label_title = tk.Label(root, text="Book Title:")
label_title.grid(row=0, column=0, padx=10, pady=5)

entry_title = tk.Entry(root)
entry_title.grid(row=0, column=1, padx=10, pady=5)

label_author = tk.Label(root, text="Author:")
label_author.grid(row=1, column=0, padx=10, pady=5)

entry_author = tk.Entry(root)
entry_author.grid(row=1, column=1, padx=10, pady=5)

label_isbn = tk.Label(root, text="ISBN:")
label_isbn.grid(row=2, column=0, padx=10, pady=5)

entry_isbn = tk.Entry(root)
entry_isbn.grid(row=2, column=1, padx=10, pady=5)

#This button will add the book info entered by the user
button_add_book = tk.Button(root, text="Add Book", command=add_book)
button_add_book.grid(row=3, column=0, columnspan=2, pady=10)

if __name__ == "__main__":
    create_gui()