from tkinter import *
import tkinter.font as font
import pandas as pd
import smtplib

def hans(message):
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login("hackathonatndc@gmail.com", "dbprqyvzrnwklxpx")

    # sending the mail
    s.sendmail ("hackathonatndc@gmail.com", a, message)

    # terminating the session
    s.quit()

#Read spreadsheet
df = pd.read_excel(r"C:\Users\jsury\Desktop\detail.xlsx")

#Select data in column
attend = (pd.DataFrame(df, columns=["Attendance", "Email", "Name"]))

#loop to get individual attendancce data
for ind in df.index:
    if df['Attendance'][ind] == 75:
        a = df['Email'][ind]
        message = 'Subject: {}\n\n{}'.format("Regarding attendance", "Dear student,\nYour attendance is 75% and you must maintain this if you want to be eligible for exam")
        hans(message)
    elif df['Attendance'][ind] < 75:
        a = df['Email'][ind]
        message = 'Subject: {}\n\n{}'.format("Regarding attendance", "Dear student,\nYour attendance is less than 75% and you will not be able to attend exam")
        hans(message)
    elif df['Attendance'][ind] <= 80:
        a = df['Email'][ind]
        message = 'Subject: {}\n\n{}'.format("Regarding attendance", "Dear student,\nYour attendance is less than 80%")
        hans(message)

gui = Tk(className='Python Examples - Button')
gui.geometry("500x200")

# define font
myFont = font.Font(family='Helvetica', size=20, weight='bold')
label = Label(gui, text="Success", font=("Courier", 44))
label.pack()

gui.mainloop()