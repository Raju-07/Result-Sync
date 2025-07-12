from customtkinter import *

app = CTk()
app.geometry("888x888")
 


lable = CTkLabel(master=app,height=100,width=55,text="Result Sync",text_color="#4CAF50",compound=CENTER,font=("Roboto",16,"bold"),image=None)
lable.place(x=0.5,y=5)

button = CTkButton(master=app,width=150,height=35,corner_radius=8,text="Click me",font=("Roboto",16,"normal"),fg_color="#3E9B3E",hover_color='green',)
button.place(relx=0.2,rely=0.2)

bu = CTkButton(master=app)
bu.place(relx=0.3,rely=0.4)


entry = CTkEntry(master=app)
entry.place(relx=0.8,rely=0.2)

entry2 = CTkEntry(master=app,width=180,height=35,corner_radius=7,border_color="#4CAF50",border_width=2,placeholder_text="Input Record File",font=('Roboto',14,'normal'))
entry2.place(relx=0.7 , rely = 0.4)


app.mainloop()