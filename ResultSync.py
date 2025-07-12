from customtkinter import *

class AppResult:

    def __init__(self):
        # setting custom Appearance
        set_default_color_theme("green")
        set_appearance_mode('system')

        # Creating Window
        app = CTk()
        app.geometry('1000x550')
        app.title("Result Sync")
        

        app.mainloop()


app = AppResult()