from tkinter import messagebox
import tkinter as tk
import pymongo
import pandas as pd

myclient = pymongo.MongoClient("mongodb://127.0.0.1:27017/")
mydb = myclient["bookstorebase"]
mycol = mydb["books"]

lst = [['Book ID', 'Title', 'Page', 'Year', 'Author']]

def export_to_excel():
    # Create a DataFrame from the grid data
    df = pd.DataFrame(lst[1:], columns=lst[0])

    # Choose a file name for the Excel file
    excel_file = "book_data.xlsx"

    # Export the DataFrame to Excel
    df.to_excel(excel_file, index=False)

    # Show a message box indicating the export success
    messagebox.showinfo("Export Successful", f"Data exported to {excel_file}")
def callback(event):
    li = []
    li = event.widget._values
    cId.set(lst[li[1]][0])
    cTitle.set(lst[li[1]][1])
    cPage.set(lst[li[1]][2])
    cYear.set(lst[li[1]][3])
    cAuthor.set(lst[li[1]][4])

def creategrid(n):
    lst.clear()
    lst.append(['Book ID', 'Title', 'Page', 'Year', 'Author'])
    cursor = mycol.find({})
    for data in cursor:
        bookId = str(data['bookId'])
        bookTitle = str(data['bookTitle'].encode('utf-8').decode('utf-8'))
        bookPage = str(data['bookPage'])
        bookYear = str(data['bookYear'])
        author = str(data['author'].encode('utf-8').decode('utf-8'))
        lst.append([bookId, bookTitle, bookPage, bookYear, author])

    for i in range(len(lst)):
        for j in range(len(lst[0])):
            mgrid = tk.Entry(window, width=30)
            mgrid.insert(tk.END, lst[i][j])
            mgrid._values = mgrid.get(), i
            mgrid.grid(row=i + 5, column=j + 1)
            mgrid.bind("<Button-1>", callback)

    if n == 1:
        for label in window.grid_slaves():

            if int(label.grid_info()["row"]) > 7:
                label.grid_forget()
def msg(msg,titlebar):
    result = messagebox.askokcancel(title=titlebar, message=msg)
    return result

def save():
    r = msg("save record?", "save")
    if r == True:
        newid = mycol.count_documents({})
        if newid!=0:
            newid = mycol.find_one(sort=[("bookId", -1)])["bookId"]
        id = newid + 1
        cId.set(id)
        mydict = { "bookId": int(custId.get()),
                   "bookTitle": custTitle.get(),
                   "bookPage": int(custPage.get()),
                   "bookYear": int(custYear.get()),
                   "author": custAuthor.get()
                    }
        x = mycol.insert_one(mydict)

        custTitle.delete(0, 'end')
        custPage.delete(0, 'end')
        custYear.delete(0, 'end')
        custAuthor.delete(0, 'end')

        creategrid(1)
        creategrid(0)


def delete():
    r = msg("delete record?", "delete")
    if r== True:
        myquery = {"bookId": int(custId.get())}
        mycol.delete_one(myquery)

        creategrid(1)
        creategrid(0)
def update():
    r = msg("update record", "update")
    if r == True:

        myquery = {"bookId": int(custId.get())}
        newvalues = {"$set": {"bookTitle": custTitle.get()}}
        mycol.update_one(myquery, newvalues)

        myquery = {"bookId": int(custId.get())}
        newvalues = {"$set": {"bookPage": custPage.get()}}
        mycol.update_one(myquery, newvalues)

        myquery = {"bookId": int(custId.get())}
        newvalues = {"$set": {"bookYear": custYear.get()}}
        mycol.update_one(myquery, newvalues)

        myquery = {"bookId": int(custId.get())}
        newvalues = {"$set": {"author": custAuthor.get()}}
        mycol.update_one(myquery, newvalues)

        creategrid(1)
        creategrid(0)

window = tk.Tk()
window.title("Book Management System")
window.geometry("1100x500")
window.configure(bg="gray")



label = tk.Label(window, text="Book ID:", width=20, height=2, bg="gray")
label.grid(column=1, row=2)
cId = tk.StringVar()
custId = tk.Entry(window, textvariable=cId, width=30)
custId.grid(column=1, row=3)
custId.configure(state=tk.DISABLED)

label = tk.Label(window, text="Book title:", width=20, height=2, bg="gray")
label.grid(column=2, row=2)
cTitle = tk.StringVar()
custTitle = tk.Entry(window, textvariable=cTitle, width=30)
custTitle.grid(column=2, row=3)

label = tk.Label(window, text="Book page:", width=20, height=2,bg="gray")
label.grid(column=3, row=2)
cPage = tk.StringVar()
custPage = tk.Entry(window, textvariable=cPage, width=30)
custPage.grid(column=3, row=3)

label = tk.Label(window, text="Book year:", width=20, height=2,bg="gray")
label.grid(column=4, row=2)
cYear = tk.StringVar()
custYear = tk.Entry(window, textvariable=cYear, width=30)
custYear.grid(column=4, row=3)

label = tk.Label(window, text="Author:", width=20,bg="gray")
label.grid(column=5, row=2)
cAuthor = tk.StringVar()
custAuthor = tk.Entry(window, textvariable=cAuthor, width=30)
custAuthor.grid(column=5, row=3)

label = tk.Label(window, text="Book Management System",bg="gray",height=2, anchor="center")
label.config(font=("Courier", 10))
label.grid(column=3, row=4)

creategrid(0)


saveBtn = tk.Button(text="Save", command=save, width=10)
saveBtn.grid(column=6, row=2, padx=(20, 0))
deleteBtn = tk.Button(text="Delete", command=delete, width=10)
deleteBtn.grid(column=6, row=3, columnspan=2, padx=(20, 0))
updateBtn = tk.Button(text="Update", command=update, width=10)
updateBtn.grid(column=6, row=4, columnspan=2, padx=(20, 0))
exportBtn = tk.Button(text="Export to Excel", command=export_to_excel)
exportBtn.grid(column=6, row=5, columnspan=2, padx=(20, 0))

window.mainloop()