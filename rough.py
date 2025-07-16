import time
import random
# from customtkinter import *

# app = CTk()
# app.geometry("888x888")
 


# lable = CTkLabel(master=app,height=100,width=55,text="Result Sync",text_color="#4CAF50",compound=CENTER,font=("Roboto",16,"bold"),image=None)
# lable.place(x=0.5,y=5)

# button = CTkButton(master=app,width=150,height=35,corner_radius=8,text="Click me",font=("Roboto",16,"normal"),fg_color="#3E9B3E",hover_color='green',)
# button.place(relx=0.2,rely=0.2)

# bu = CTkButton(master=app)
# bu.place(relx=0.3,rely=0.4)


# entry = CTkEntry(master=app)
# entry.place(relx=0.8,rely=0.2)

# entry2 = CTkEntry(master=app,width=180,height=35,corner_radius=7,border_color="#4CAF50",border_width=2,placeholder_text="Input Record File",font=('Roboto',14,'normal'))
# entry2.place(relx=0.7 , rely = 0.4)


# app.mainloop()


# # Method 1: Using round()
# num = 3.14159
# print("Using round():", round(num, 2))

# # Method 2: Using string formatting
# print("Using format(): {:.2f}".format(num))

# # Method 3: Using f-string (Python 3.6+)
# print(f"Using f-string: {num:.2f}")

"Logic for Progressbar"
total_record = random.randint(1,100)
one_time_inc = 1/total_record

current_inc = 0
for _ in range(total_record):
    current_inc += one_time_inc
#    print(f"Progress bar status {current_inc:.2f}")
    print(f"Percentage {current_inc*100:.2f}%")


"Logic for Percentage"
# total_record = 65
# one_time_percentage = 1/total_record
# print(one_time_percentage)
# current_percentage = 0.0
# for i in range(1,total_record+1):
#     current_percentage = (i/total_record)*100
#     print(f"Current Percentage: {current_percentage:.2f}%")

"Logic for Estimate tiime"

# for i in range(10000000):
#     start_time = time.time()
#     print(start_time)
#     time.sleep(random.randint(1,5))
#     end_time = time.time()
#     print(f"{10000000/(end_time-start_time)}")

# for i in range(1,65+1):
#     start_time = time.time()
#     time.sleep(4)
#     end_time = time.time()
#     print((65-i)/(end_time-start_time))
