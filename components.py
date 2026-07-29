from customtkinter import CTkLabel,CTkButton,CTkEntry,CENTER,LEFT

class ReusableComponents:
    def create_label(self, master, text:str, x:float, y:float, font=("Roboto", 18, "normal"), text_color=None):
        label = CTkLabel(master=master, text=text, font=font, text_color=text_color)
        label.place(relx=x, rely=y, anchor=CENTER)
        return label
    
    def create_button(self, master, text:str, relx:float, rely:float, command, width=150, height=35, font=("Roboto", 16, "normal"), image=None, compound=LEFT, fg_color="#22C55E", hover_color="#16A34A"):
        button = CTkButton(master=master, width=width, height=height, corner_radius=8, font=font, image=image, compound=compound, text=text, command=command, fg_color=fg_color, hover_color=hover_color)
        button.place(relx=relx, rely=rely, anchor=CENTER)
        return button
    
    def create_entrybox(self, master, placeholder_text, relx:float, rely:float, width=180, height=35, cor_rad=6, bdr_clr='#1F2937', bdr_wid=2, font=("Roboto", 14, "normal")):
        entrybox = CTkEntry(master=master, placeholder_text=placeholder_text, corner_radius=cor_rad, width=width, height=height, border_color=bdr_clr, border_width=bdr_wid, font=font, fg_color="#020617", text_color="#E5E7EB")
        entrybox.place(relx=relx, rely=rely, anchor=CENTER)
        return entrybox