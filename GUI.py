import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from Data_Processing import data_processing
from Email_Sending import email_sending

# The layout of first window
window1 = tk.Tk()
window1.title('Data Processing')
window1.geometry('500x500')

lab1 = tk.Label(window1, text='Please input the sheet number you want to use:')
lab1.place(x=0, y=25)
lab2 = tk.Label(window1, text='Sheet in use:')
lab2.place(x=0, y=75)
lab3 = tk.Label(window1, text='The email address going to be sent:')
lab3.place(x=0, y=125)
str1 = tk.StringVar()
lab4 = tk.Label(window1, textvariable=str1, fg='blue')
lab4.place(x=90, y=75)
str2 = tk.StringVar()
lab5 = tk.Label(window1, textvariable=str2, fg='blue')
lab5.place(x=0, y=150)
e1 = tk.Entry(window1, width=3)
e1.place(x=315, y=25)

# Set the default of the sheet number
sheet_number = 0

on_hit = False


def submit():
    global on_hit

    if not on_hit:
        on_hit = True
        sheet_number = int(e1.get()) - 1
        sheet_names, receiver = data_processing('Data.xlsx', sheet_number)
        str1.set(sheet_names[sheet_number])
        str2.set(receiver)
    else:
        on_hit = False
        str1.set('')
        str2.set('')


def finish():
    window1.destroy()

    # Open the Login window
    # The layout of the second window
    window2 = tk.Tk()
    window2.title('Login')
    window2.geometry('400x100')

    win2_lab1 = tk.Label(window2, text='Please input the password:')
    win2_lab1.place(x=0, y=25)
    win2_e1 = tk.Entry(window2, show='*')
    win2_e1.place(x=180, y=25)

    # Check the password after clicking the OK button
    def check():
        ans = win2_e1.get()
        if ans == '0000':
            window2.destroy()

            # Open the email writing window
            # The layout of the third window
            window3 = tk.Tk()
            window3.title('Email Sending')
            window3.geometry('1000x800')

            win3_lab1 = tk.Label(window3, text='Subject:')
            win3_lab1.place(x=0, y=25)
            win3_lab2 = tk.Label(window3, text='Contents:')
            win3_lab2.place(x=0, y=75)
            win3_lab3 = tk.Label(window3, text='Attachment:')
            win3_lab3.place(x=0, y=540)
            win3_str1 = tk.StringVar()
            win3_lab4 = tk.Label(window3, textvariable=win3_str1, bg='white')
            win3_lab4.place(x=85, y=542)
            win3_e1 = tk.Entry(window3, width=97)
            win3_e1.place(x=85, y=23)
            win3_e2 = tk.Text(window3, width=125, height=30)
            win3_e2.place(x=85, y=73)

            # Get the file path when user hit the attachment button
            def get_file_path():
                file_path = filedialog.askopenfilename()
                win3_str1.set(file_path)

            # Send the email after clicking the Send button
            def send():
                subject = win3_e1.get()
                contents = win3_e2.get('1.0', 'end')
                email_sending(subject, contents, sheet_number, attachment=win3_str1.get())
                window3.destroy()

            win3_b1 = tk.Button(window3, text='Send', command=send)
            win3_b1.place(x=479, y=775)
            win3_b2 = tk.Button(window3, text='Add attachment', command=get_file_path)
            win3_b2.place(x=85, y=542)
            window3.mainloop()

        else:
            tk.messagebox.showerror(message='Wrong password. Try again.')

    win2_b1 = tk.Button(window2, text='OK', command=check)
    win2_b1.place(x=186, y=75)
    window2.mainloop()


b1 = tk.Button(window1, text='Submit', command=submit)
b1.place(x=355, y=27)
b2 = tk.Button(window1, text='OK', command=finish)
b2.place(x=236, y=475)
window1.mainloop()
