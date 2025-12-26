import webbrowser
from customtkinter import *

app = CTk()
app.geometry("444x444")
app.title("Improvement ")

def on_click():
   webbrowser.open("https://forms.gle/PoXZRNdXS1P14MsaA")

label = CTkLabel(app,text="Help me to make this program more efficient")
label.pack()
labelhelp = CTkLabel(app,text="Click here")
labelhelp.pack()
labelhelp.bind("Button-1",lambda event: on_click())

# Bind left mouse click to the label
label.bind("<Button-1>", lambda event: on_click())

# Optional: Change appearance on hover
label.bind("<Enter>", lambda event: label.configure(font=("Arial", 14, "underline"), cursor="hand2"))
label.bind("<Leave>", lambda event: label.configure(font=("Arial", 12), cursor="arrow"))

app.mainloop()
app.mainloop()

