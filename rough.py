from customtkinter import *
from tkinter import messagebox as msg
import os

def createfiles():
    new_files = ["Processed_entries.txt","row_status.txt"]
    try:
        for files in new_files:
            if not os.path.exists(files):
                with open(files,'w') as file:
                    file.close()
        msg.showinfo("All the files are created")
    except Exception as e:
        msg.showerror("Error occur during file creation")
        print(f"Reason is {e}")

def showmessage():
    cache_file = ["Processed_entries.txt","row_status.txt"]
    try:
        if msg.askyesno("Cache Deletion Confirmation", "Would you like to proceed with deleting the file?"):
            for files in cache_file:
                if os.path.exists(files):
                    os.remove(files)
            msg.showinfo("All the files has been sucessfully deleted..")
        else:
            print("File deletion was not executed.")
    except Exception as e:
        print("The cache deletion process was interrupted.")
        print(f"Error: {e}")

app = CTk()
app.geometry("444x555")
app.title("Hello")

# button
button1 = CTkButton(master=app, width=120, height=45, corner_radius=5, text="Delete files", bg_color="green", command=showmessage)
button1.pack(pady=20)  # Use pack to place the button with padding
button2 = CTkButton(master=app, width=120, height=45, corner_radius=5, text="Create files", bg_color="green", command=createfiles)
button2.pack(pady=20)  # Use pack to place the button with padding

app.mainloop()