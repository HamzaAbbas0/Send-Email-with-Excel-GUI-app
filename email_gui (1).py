import smtplib, ssl
import sys
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from smtplib import SMTPException
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
# from mailer import Mailer
import os

my_w = Tk()
my_w.geometry("400x580")
my_w.maxsize(400 ,580)
my_w.title("Mass Email")
my_font1 = ("times 18 bold")
photo = PhotoImage(file="Group 982.png")
my_w.iconphoto(False, photo)


def exit_function():
    my_w.destroy()
    os._exit(1)


### Function to send the email ###
def send_an_email():
    global excel_file
    print(excel_file)
    data1 = e1.get()
    print(data1)
    data2 = e2.get()
    # print(data2)
    data3 = inputtxt.get(1.0, "end-1c")
    # print(data3)
    data4 = e3.get()
    me = data2
    subject = data4
    body = data3
    print(me)
    print(subject)
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = me
    msg.preamble = body
    msg.attach(MIMEText(body))

    #part = MIMEBase('application', "octet-stream")
    # part.set_payload(open("1.png", "rb").read())
    # encoders.encode_base64(part)
    # part.add_header('Content-Disposition')
    # print('emAIL SENT')
    #msg.attach(part)

    try:
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.ehlo()
        s.starttls()
        s.ehlo()
        # s.login(user='hackerspacekarachi@gmail.com', password='lahore@@1')
        print("login")
        s.login(user=data1, password=data2)
        # s.send_message(msg)
        email_list = pd.read_excel(excel_file)
        # getting the names and the emails
        names = email_list['NAME']
        emails = email_list['EMAIL']
        print("email list", range(len(emails)))
        #emails = ['iotwork40@gmail.com']
        for i in range(len(emails)):
            # for every record get the name and the email addresses
            name = names[i]
            email = emails[i]
            print(names[i])
            # print("i=", i)
            # toaddr = [email]
            #email = emails[i]
            # s.sendmail(data2, email, data3,msg.as_string(), data4)
            s.sendmail(me, email, msg.as_string(), subject)
        messagebox.showinfo("Mass Email", "Your Message is Send Successfully")

        s.quit()
        
    # except:
    #   print ("Error: unable to send email")
    except SMTPException as error:
        print("Error")
        messagebox.showerror("Mass Email", "Please Type correct email and password")


def upload_file():
    global excel_file
    file = filedialog.askopenfilename(
        filetypes=[("Excel files", ".xlsx")]
    )
    excel_file = file
    # email_list = pd.read_excel(file)
    # print(file)
    if (file):
        my_str.set(file)
    # fob = open(file,"r")
    # #     print(fob.read())


b1 = Button(my_w, text="Upload file", bd=4, command=lambda: upload_file())
b1.place(x=230, y=440)
# b2 = Button(my_w, text="upload files",)
# b2.place(x=300,y=300)
b3 = Button(my_w, text="Send", width=8, bd=4, command=send_an_email)
b3.place(x=320, y=440)

my_str = StringVar()
l2 = Label(my_w, textvariable=my_str, fg="red")
l2.place(x=60, y=500)
my_str.set("Please select the Excel file from the Upload file button")

e1 = Entry(my_w, bd=5, relief=SUNKEN, width=40, )
e1.place(x=120, y=40)
e2 = Entry(my_w, bd=5, relief=SUNKEN, width=40,show="*" )
e2.place(x=120, y=90)
e3 = Entry(my_w, bd=5, relief=SUNKEN, width=40, )
e3.place(x=120, y=150)

l3 = Label(my_w, text="Email address :", font=("Arial 9 bold"))
l3.place(x=10, y=44)

l4 = Label(my_w, text="Password :", font=("Arial 9 bold"))
l4.place(x=10, y=94)
l5 = Label(my_w, text="Message", font=("Arial 9 bold"))
l5.place(x=10, y=224)
l6 = Label(my_w, text="Subject :", font=("Arial 9 bold"))
l6.place(x=10, y=155)

inputtxt = Text(my_w, height=8, width=47)
inputtxt.place(x=10, y=250)
my_w.protocol("Delete_window", exit_function)

my_w.mainloop()
# send_an_email()
