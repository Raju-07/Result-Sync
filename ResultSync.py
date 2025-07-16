# Required Library for the Project
#Module for GUI
from customtkinter import *
from tkinter import messagebox as msg
from tkinter import END
#Module for Automating Excel data Read and Write
from openpyxl import Workbook,load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill,Font,Alignment

#Automate the task using selenium"
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
        self.app.configure(fg_color = "#F5F7FA")

        # ICON
        self.file_path = os.path.dirname(os.path.realpath(__file__))        
        self.folder_icon = CTkImage(Image.open(self.file_path + "/drive_file_icon.png"),size=(33,33))


        #New Frame
        Frame1 = CTkFrame(master=self.app,width=310,height=450,fg_color="#E3E9F1",corner_radius=10,border_color='#CBD5E1',border_width=2)
        Frame1.place(relx=0.83,rely=0.55,anchor=CENTER)
        
        # Adding Labels to the self.application
        self.Label(self.app,"Result Sync",0.5,0.06,font=("Eras Bold ITC",50,"bold"),text_color="#1F2937")     #Heading
        self.Label(self.app,"━━━━ Automate Result Extraction & Excel Reporting ━━━━",0.5,0.13,text_color="#475569")   #Subheading
        self.Label(self.app,"©2025 Raju Yadav | Powered by Python made with Love ❤",0.5,0.97,text_color="#64748B")   #Footer

        self.Label(self.app,"Status",0.4,0.33,("Eras Bold ITC",22,"bold"))
        self.current_student = self.Label(self.app,"Not Started Yet",0.4,0.37,("Roboto",18,"normal"),text_color="#1ACCF8",)

        # Frame1
        self.Label(Frame1,"Required Info",0.5,0.07,font=("Eras Bold ITC",27,"bold")) # Frame Heading

        self.Label(Frame1,"Select Record File (.xlsx)",0.33,0.2,font=("Roboto",16,"normal")) # input Excel Record File
        self.input_file = self.Entrybox(Frame1,"Input Record File Path",0.42,0.28,230)    # For File Path
        self.Button(Frame1,"",0.9,0.28,self.SelectRecordFile,font=("Roboto",16,'normal'),width=10,height=10,image=self.folder_icon,fg_color="#E5E7EB",hover_color="#10B981")

        self.Label(Frame1,"Location to Save",0.26,0.4,font=("Roboto",16,"normal")) # choose Location and Name for Output File
        self.output_file = self.Entrybox(Frame1,"Select Where to Save",0.42,0.48,230) # For Location and Name of the File
        self.Button(Frame1,"",0.9,0.48,self.selectLocName,font=("Roboto",16,'normal'),width=10,height=10,image=self.folder_icon,fg_color="#E5E7EB",hover_color="#10B981")
        
        self.Button(Frame1,"Proceed",0.5,0.65,self.Proceed,font=("Roboto",18,"bold"))   #For Processing
        # Important Note
        self.Label(Frame1,"NOTE: FIRST ROW IS TREATED AS HEADERS REGISTRATION\n NO. MUST BE IN COLUMN A & ROLL NO. IN COLUMN B",0.5,0.8,font=("calibri",12,"normal"),text_color='#EF4444')
        self.Label(Frame1,"NOTE: IF AN ERROR OCCURS, RESELECT THE SAME OUTPUT\n FILE ",0.5,0.9,font=("calibri",12,"normal"),text_color='#EF4444')



        self.app.mainloop()
    
    # Create Lable
    def Label(self,master,text:str,x:float,y:float,font = ("Roboto",18,"normal"),text_color=None):
        label = CTkLabel(master=master,text=text,font=font,text_color=text_color)
        label.place(relx=x,rely=y,anchor=CENTER)
        return label
    
    # Create Button
    def Button(self,master,text:str,relx:float,rely:float,command,width=150,height=35,font=("Roboto",16,"normal"),image=None,compound=LEFT,fg_color="#10B981",hover_color="#059669"):
        button = CTkButton(master=master,width=width,height=height,corner_radius=8,font=font,image=image,compound=compound,text=text,command=command,fg_color=fg_color,hover_color=hover_color)
        button.place(relx=relx,rely=rely,anchor=CENTER)
        return button
    
    # Create Entrybox
    def Entrybox(self,master,placeholder_text,relx:float,rely:float,width=180,height=35,cor_rad=6,bdr_clr='#CBD5E1',bdr_wid=2,font=("Roboto",14,"normal")):
        entrybox = CTkEntry(master=master,placeholder_text=placeholder_text,corner_radius=cor_rad,width=width,height=height,border_color=bdr_clr,border_width=bdr_wid,font=font,fg_color="#FFFFFF")
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
        return row_count

    def Proceed(self):
        if self.input_file.get() and self.output_file.get():
            self.input_file.configure(border_color='green')
            self.output_file.configure(border_color = "green")

            self.estimate_time = self.TotalRecord()

            #Progress bar
            self.progress_status = 0.0
            self.percentage_status = self.progress_status*100

            self.progressbar = CTkProgressBar(self.app,300,height=18,corner_radius=6,border_color="blue",progress_color="#34D399",orientation="horizontal")
            self.progressbar.place(relx=0.42,rely=0.5,anchor=CENTER)
            self.progressbar.set(self.progress_status) 
            self.percentage_label = self.Label(self.app,f"{self.percentage_status}%",0.6,0.5)

            self.est_time_lable = self.Label(self.app,f"Estimate Time: {self.estimate_time} Minutes",0.43,0.55)

            self.start_btn = self.Button(self.app,"Start",0.41,0.65,command=self.LetsDoIt)

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
    #initializing a file which track the of the number of row for automate excel data
    def row_tracker(self):
        self.current_row = 4
        self.row_tracker_file = "row_status.txt"
        if os.path.exists(self.row_tracker_file):
            with open(self.row_tracker_file,'r') as file:
                row = file.readline().strip()
        else:
            with open(self.row_tracker_file,'w') as file:
                file.write(str(self.current_row))
            row = str(self.current_row)
        return row


    def LetsDoIt(self):
        #Getting the File in which student data is stored
        if self.input_file.get():
            input_wb = load_workbook(self.input_file.get())
            input_ws = input_wb.active
        else:
            msg.showerror("File Not Loaded","Error during Loading Input File")
            exit()
        #Fetching last row
        self.current_row = int(self.row_tracker())
        #one time increament for the progressbar
        self.one_time_increment = 1/self.TotalRecord()

        #initializing Excel File Here
        if os.path.exists(self.output_file.get()):
            wb = load_workbook(self.output_file.get())
        else:
            wb = Workbook()

        if "Sheet1" in wb.sheetnames:
            ws = wb['Sheet1']
        else:
            ws = wb.active
            ws.title = "Sheet1"

        # --- Title Header (Row 1) ---
        ws.merge_cells('A1:N1')
        ws['A1'].fill = PatternFill(start_color="8DFAB1",end_color="8DFAB1",fill_type="solid")
        ws['A1'] = "BCA Result 2025 Sem - IV"
        ws['A1'].font = Font("Times New Roman",size="16",bold=True,color="006100")
        ws['A1'].alignment = Alignment(horizontal="center",vertical="center")

        # --- Section Headers (Row 2) ---
        section_headers = [
            {"cell": "A2:E2", "text": "Students Details"},
            {"cell": "G2:K2", "text": "Marks in Each Subjects"},
            {"cell": "M2:N2", "text": "Result"}]

        for section in section_headers:
            ws.merge_cells(section["cell"])
            cell_ref = section["cell"].split(":")[0]  # Get the top-left cell of the merged area
            ws[cell_ref] = section["text"]
            ws[cell_ref].fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
            ws[cell_ref].font = Font(name="Century", size=14, color="006100")
            ws[cell_ref].alignment = Alignment(horizontal="center", vertical="center")


        # Adding Header of the Columns Formatting
        col_head = ["Sr.No","Name","Father Name","Registration No","Roll No","","BCA206","BCA207","BCA208","BCA209","BCA210","","Total Marks","Result"]

        cols = 1
        for header in col_head:
            cols_char = get_column_letter(cols)
            ws[cols_char+str(3)] = header
            cols += 1


        # Set up the WebDriver (Edge in this example)
        options = Options()
        options.add_experimental_option('excludeSwitches',['enable-logging'])  # Suppress DevTools message
        driver = webdriver.Edge(options=options)
        wait = WebDriverWait(driver,10)
        # Navigate to the result portal
        driver.get("https://result.mdu.ac.in/postexam/result.aspx")

        #Updating Current row
        self.current_row = int(self.row_tracker())
        self.process_data = []
        file_name = "Processed_entries.txt"
        try:
            if os.path.exists(file_name):
                with open(file_name,"r") as file:
                    self.process_data = list(map(int,file.read().split()))
        except Exception as e:
            msg.showerror("Error","There is an error While reading the Processed Entries ")

        finally:
            with open(file_name,"a") as file:
                for rows in input_ws.iter_rows(min_row=2,max_col=2,values_only=True):
                    try:
                        if rows[0] is None:
                            break
                        elif rows[0] in self.process_data:
                            continue
                        else:
                            # Process of Extracting Result
                            driver.find_element(By.ID, "txtRegistrationNo").send_keys(rows[0])  # Registration Number one by one
                            driver.find_element(By.ID, "txtRollNo").send_keys(rows[1])  # Roll Number one by one
                            # Submit the form
                            wait.until(EC.element_to_be_clickable((By.ID,"cmdbtnProceed"))).click()
                            wait.until(EC.element_to_be_clickable((By.ID,"imgComfirm"))).click()
                            wait.until(EC.element_to_be_clickable((By.ID,"rptMain_ctl01_lnkView"))).click()
                            wait.until(EC.visibility_of_element_located((By.ID,"lblStudentName")))
                            student_name = driver.find_element(By.ID,"lblStudentName").text
                            father_name = driver.find_element(By.ID,"lblFatherName").text
                            #Just for Checking Current status
                            self.current_student.configure(text=student_name)

                            #Extracting Subjects Marks
                            subject_marks = [driver.find_element(By.XPATH, f"//table//tr[{i}]//td[8]").text for i in range(2, 7)]
                        #random_data = driver.find_element(By.XPATH("//table//tr[2]//td[8]"))#just for example a simple way for extracting random data
                            bca201, bca202, bca203, bca204, bca205 = subject_marks
                            
                            #Result
                            total_marks = driver.find_element(By.ID,"rptMarks_ctl06_lblTotal").text
                            result = driver.find_element(By.ID,"lblresult").text
                    
                    # Storing data in Excel for each row
                            data = [self.current_row-3,student_name,father_name,rows[0],rows[1],"",bca201,bca202,bca203,bca204,bca205,"",total_marks,result]
                            for char in range(1,15):
                                char_num = get_column_letter(char)
                                ws[char_num+str(self.current_row)] = data[char-1]
                                with open(self.row_tracker_file,"w") as file2:
                                    file2.write(str(self.current_row))
                            self.current_row += 1
                            

                        #Updating Progressbar , Percentage and Estimate Time

                            self.progressbar.set(self.progress_status)
                            self.progress_status += self.one_time_increment

                            self.percentage_label.configure(text=f"{(self.progress_status)*100:.2f}%")

                            self.est_time_lable.configure(text=f"Estimate Time: {self.estimate_time/4} Minutes")

                            self.estimate_time -= 1


                            #Updating registraion_no in processed_entries file
                            file.write(str(rows[0]) + "\n")
                            driver.refresh() #Reloading website for repreating process

                    except Exception as e:
                        wb.save(self.output_file.get())
                        print(f"Error is occured In {rows[0]} and \nUnexpected Error: Please ensure good internet connect or try again")
                        msg.showerror("process Interupted",f"{e}")
                        exit()
                            
                wb.save(self.output_file.get())        
        #updating the last successful record row
            with open(self.row_tracker_file,'w') as track_update:
                track_update.write(str(self.current_row))
            # Display info when process complete
            msg.showinfo("Success","All the data fatched successfully")
            # Close the browser
            driver.quit()





app = AppResult()