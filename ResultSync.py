from customtkinter import *
from tkinter import messagebox as msg
from openpyxl import Workbook,load_workbook
import os
from PIL import Image,ImageTk


class AppResult:

    def __init__(self):
        # setting custom self.appearance
        set_default_color_theme("green")
        set_appearance_mode('system')

        # Creating Window
        self.app = CTk()
        self.app.geometry('1000x600')
        self.app.title("Result Sync")

        # ICON
        self.file_path = os.path.dirname(os.path.realpath(__file__))        
        self.folder_icon = CTkImage(Image.open(self.file_path + "/icon.png"),size=(33,33))


        #New Frame
        Frame1 = CTkFrame(master=self.app,width=310,height=450,corner_radius=9,border_color='#1E293B',border_width=2)
        Frame1.place(relx=0.83,rely=0.55,anchor=CENTER)
        
        # Adding Labels to the self.application
        self.Label(self.app,"Result Sync",0.5,0.06,font=("Eras Bold ITC",50,"bold"))     #Heading
        self.Label(self.app,"━━━━ Automate Result Extraction & Excel Reporting ━━━━",0.5,0.13)   #Subheading
        self.Label(self.app,"©2025 Raju Yadav | Powered by Python made with Love ❤",0.5,0.97)   #Footer

        self.Label(self.app,"Status",0.4,0.33,("Eras Bold ITC",22,"bold"))
        self.Label(self.app,"Not Started Yet",0.4,0.37,("Roboto",18,"normal"),text_color="#1ACCF8",)

        # Frame1
        self.Label(Frame1,"Required Info",0.5,0.07,font=("Eras Bold ITC",27,"bold")) # Frame Heading

        self.Label(Frame1,"Select Record File (.xlsx)",0.33,0.2,font=("Roboto",16,"normal")) # input Excel Record File
        self.input_file = self.Entrybox(Frame1,"Input Record File Path",0.42,0.28,230)    # For File Path
        self.Button(Frame1,"",0.9,0.28,self.SelectRecordFile,font=("Roboto",16,'normal'),width=10,height=10,image=self.folder_icon,fg_color="#dbdbdb",hover_color="#f0eeee")

        self.Label(Frame1,"Location to Save",0.26,0.4,font=("Roboto",16,"normal")) # choose Location and Name for Output File
        self.output_file = self.Entrybox(Frame1,"Select Where to Save",0.42,0.48,230) # For Location and Name of the File
        self.Button(Frame1,"",0.9,0.48,self.selectLocName,font=("Roboto",16,'normal'),width=10,height=10,image=self.folder_icon,fg_color="#dbdbdb",hover_color="#f0eeee")
        
        self.Button(Frame1,"Proceed",0.5,0.65,self.Proceed,font=("Roboto",18,"bold"))   #For Processing
        # Important Note
        self.Label(Frame1,"NOTE: FIRST ROW IS TREATED AS HEADERS REGISTRATION\n NO. MUST BE IN COLUMN A & ROLL NO. IN COLUMN B",0.5,0.8,font=("calibri",12,"normal"),text_color='red')
        self.Label(Frame1,"NOTE: IF AN ERROR OCCURS, RESELECT THE SAME OUTPUT\n FILE ",0.5,0.9,font=("calibri",12,"normal"),text_color='red')



        self.app.mainloop()
    
    # Create Lable
    def Label(self,master,text:str,x:float,y:float,font = ("Roboto",18,"normal"),text_color=None):
        label = CTkLabel(master=master,text=text,font=font,text_color=text_color)
        label.place(relx=x,rely=y,anchor=CENTER)
        return label
    
    # Create Button
    def Button(self,master,text:str,relx:float,rely:float,command,width=150,height=35,font=("Roboto",16,"normal"),image=None,compound=LEFT,fg_color=None,hover_color=None):
        button = CTkButton(master=master,width=width,height=height,corner_radius=8,font=font,image=image,compound=compound,text=text,command=command,fg_color=fg_color,hover_color=hover_color)
        button.place(relx=relx,rely=rely,anchor=CENTER)
        return button
    
    # Create Entrybox
    def Entrybox(self,master,placeholder_text,relx:float,rely:float,width=180,height=35,cor_rad=6,bdr_clr='#1E293B',bdr_wid=2,font=("Roboto",14,"normal")):
        entrybox = CTkEntry(master=master,placeholder_text=placeholder_text,corner_radius=cor_rad,width=width,height=height,border_color=bdr_clr,border_width=bdr_wid,font=font)
        entrybox.place(relx=relx,rely=rely,anchor=CENTER)
        return entrybox
    
    # Tracking Total Record
    def TotalRecord(self,):
        if os.path.exists(self.input_file.get()):
            wb = load_workbook(self.input_file.get())
            ws = wb['Sheet1']
            row_count = 0
            for rows in ws.iter_rows(min_row=2,max_col=1,values_only=True):
                if rows:
                    row_count += 1
                else:
                    continue
        else:
            msg.showerror("File Not Exists","Please Provide the correct path and Location")
        print(row_count)
        return row_count

    def Proceed(self):
        if self.input_file.get() and self.output_file.get():
            self.input_file.configure(border_color='green')
            self.output_file.configure(border_color = "green")

            self.percentage = 0
            self.estimate_time = self.TotalRecord()/5 
            #Progress bar
            self.progressbar = CTkProgressBar(self.app,300,height=18,corner_radius=6,border_color="blue",progress_color="green",orientation="horizontal")
            self.progressbar.place(relx=0.43,rely=0.5,anchor=CENTER)
            self.progressbar.set(0) 
            self.percentage_label = self.Label(self.app,f"{self.percentage} %",0.6,0.5)

            self.est_time = self.Label(self.app,f"Estimate Time: {self.estimate_time} Minutes",0.43,0.55)

        else:
            if self.input_file.get():
                self.input_file.configure(border_color="green")
            else:
                self.input_file.configure(border_color = "red")
            
            if self.output_file.get():
                self.output_file.configure(border_color = "green")
            else:
                self.output_file.configure(border_color = "red")
            
            msg.showerror("File Not Select","PLEASE SELECT THE FILE FIRST")

    #select student result file function
    def SelectRecordFile(self):
        file_path = filedialog.askopenfilename(defaultextension=".xlsx",filetypes=[("Excel Files","*.xlsx")],title="Select Record File")
        self.input_file.configure(state="normal")
        self.input_file.delete(0,END)
        self.input_file.insert(0,file_path)
        self.input_file.configure(state="readonly")

    # Select Location and File Name to save the output file
    def selectLocName(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel Files","*.xlsx")],title="Select Where to Save Output File")
        self.output_file.configure(state="normal")
        self.output_file.delete(0,END)        
        self.output_file.insert(0,file_path)
        self.output_file.configure(state="readonly")

app = AppResult()