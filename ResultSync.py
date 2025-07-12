from customtkinter import *
import os
from PIL import Image,ImageTk


class AppResult:

    def __init__(self):
        # setting custom Appearance
        set_default_color_theme("green")
        set_appearance_mode('system')

        # Creating Window
        app = CTk()
        app.geometry('1000x600')
        app.title("Result Sync")

        # ICON
        self.file_path = os.path.dirname(os.path.realpath(__file__))        
        self.folder_icon = CTkImage(Image.open(self.file_path + "/drive_file_icon.png"),size=(30,30))


        #New Frame
        Frame1 = CTkFrame(master=app,width=310,height=450,corner_radius=9,border_color='#1E293B',border_width=2)
        Frame1.place(relx=0.83,rely=0.55,anchor=CENTER)
        
        # Adding Labels to the Application
        self.Label(app,"Result Sync",0.5,0.06,font=("Eras Bold ITC",50,"bold"))     #Heading
        self.Label(app,"━━━━ Automate Result Extraction & Excel Reporting ━━━━",0.5,0.13)   #Subheading
        self.Label(app,"©2025 Raju Yadav | Powered by Python made with Love ❤",0.5,0.97)   #Footer

        # Frame1
        self.Label(Frame1,"Required Info",0.5,0.07,font=("Eras Bold ITC",27,"bold")) # Frame Heading

        self.Label(Frame1,"Select Record File (.xlsx)",0.33,0.2,font=("Roboto",16,"normal")) # input Excel Record File
        self.input_file = self.Entrybox(Frame1,"Input Record File",0.33,0.28)    # For File Path
        self.Button(Frame1,"",0.8,0.28,self.SelectRecordFile,font=("Roboto",16,'normal'),width=10,height=18,image=self.folder_icon)

        self.Label(Frame1,"Location to Save",0.26,0.4,font=("Roboto",16,"normal")) # choose Location and Name for Output File
        self.output_file = self.Entrybox(Frame1,"Select Where to Save",0.33,0.48) # For Location and Name of the File
        self.Button(Frame1,"",0.8,0.48,self.selectLocName,font=("Roboto",16,'normal'),width=10,height=18,image=self.folder_icon)
        
        self.Button(Frame1,"Start",0.5,0.65,self.Extract_result,font=("Roboto",18,"bold"))   #For Starting the Process
        # Important Note
        self.Label(Frame1,"NOTE: FIRST ROW IS TREATED AS HEADERS REGISTRATION\n NO. MUST BE IN COLUMN A & ROLL NO. IN COLUMN B",0.5,0.8,font=("calibri",12,"normal"),text_color='red')
        self.Label(Frame1,"NOTE: IF AN ERROR OCCURS, RESELECT THE SAME OUTPUT\n FILE ",0.5,0.9,font=("calibri",12,"normal"),text_color='red')



        app.mainloop()
    
    # Create Lable
    def Label(self,master,text:str,x:float,y:float,font = ("Roboto",18,"normal"),text_color=None):
        label = CTkLabel(master=master,text=text,font=font,text_color=text_color)
        label.place(relx=x,rely=y,anchor=CENTER)
        return label
    
    # Create Button
    def Button(self,master,text:str,relx:float,rely:float,command,width=150,height=35,font=("Roboto",16,"normal"),image=None,compound=LEFT,):
        button = CTkButton(master=master,width=width,height=height,corner_radius=8,font=font,image=image,compound=compound,text=text,command=command)
        button.place(relx=relx,rely=rely,anchor=CENTER)
        return button
    
    # Create Entrybox
    def Entrybox(self,master,placeholder_text,relx:float,rely:float,width=180,height=35,cor_rad=6,bdr_clr='#1E293B',bdr_wid=2,font=("Roboto",14,"normal")):
        entrybox = CTkEntry(master=master,placeholder_text=placeholder_text,corner_radius=cor_rad,width=width,height=height,border_color=bdr_clr,border_width=bdr_wid,font=font)
        entrybox.place(relx=relx,rely=rely,anchor=CENTER)
        return entrybox

    def Extract_result(self):
        pass

    #select student result file function
    def SelectRecordFile(self):
        file_path = filedialog.askopenfilename(defaultextension=".xlsx",filetypes=[("Excel Files","*.xlsx")],title="Select Record File")
        self.input_file.delete(0,END)
        self.input_file.insert(0,file_path)

    # Select Location and File Name to save the output file
    def selectLocName(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel Files","*.xlsx")],title="Select Where to Save Output File")
        self.output_file.delete(0,END)        
        self.output_file.insert(0,file_path)

app = AppResult()