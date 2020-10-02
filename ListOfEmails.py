import imaplib
import email
import time
from tkinter import messagebox
from tkinter import *
import xlsxwriter

def getemail():
    start_time = time.time()

    usin = usernameInput.get()
    pain = passwordInput.get()

    usin = str(usin)
    pain = str(pain)

    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        imap.login(usin, pain)
    except:
        messagebox.showerror("Error", """
    Error caused by one of the following reasons:
    1. Invalid password or email address.
    2. "Allow less secure apps" is set to "OFF" in Gooogle Gmail. You need to give access 
    by allowing less secure apps.
    Follow this link to fix the problem: https://myaccount.google.com/lesssecureapps
    """)

    status, messages = imap.select("INBOX")
    messages = int(messages[0])
    address = []

    for i in range(messages, 0, -1):
        res, msg = imap.fetch(str(i), "(RFC822)")
        for information in msg:
            if isinstance(information, tuple):
                msg = email.message_from_bytes(information[1])
                from_ = msg.get("From")
                address.append(from_)

    address = list(dict.fromkeys(address))

    workbook = xlsxwriter.Workbook(r'C:\Users\Asus\Downloads\EmailList_%s.xlsx' % usin)
    worksheet = workbook.add_worksheet()
    
    worksheet.write("A1", "USERNAME")
    worksheet.write("B1", "EMAIL ADDRESS")

    y = 0
    for i in address:
        bucket = i.replace(">", "").split("<")
        worksheet.write(y+1, 0, bucket[0])
        if len(bucket) < 2:
            worksheet.write(y+1, 1, bucket[0])
        else:
            worksheet.write(y+1, 1, bucket[1])
        y += 1

    workbook.close()

    imap.close()
    imap.logout()

    executionTime = Label(root, bg="#fdfae5",
                    text="\n" + "execution time --- %s seconds ---" % (time.time() - start_time))
    executionTime.grid(row=7, column=1)

    success = Label(root, bg="green", fg="white", text="Download complete")
    success.grid(row=6, column=1)



root = Tk(className="EveryEmailAddress")
root.configure(bg="#fdfae5")

Label(root, text="""Get an Excel sheet of every email address,
        which has ever written to you. With only one click!""", bg="#fdfae5").grid(row=0, column=1)

Label(root, text="Email Address: ", bg="#fdfae5").grid(row=1, column=0)
usernameInput = Entry(root, width=40)
usernameInput.focus()
usernameInput.bind("<Return>", getemail)
usernameInput.grid(row=1, column=1)

Label(root, text="Password: ", bg="#fdfae5").grid(row=2, column=0)
passwordInput = Entry(root, width=40, show="*")
passwordInput.focus()
passwordInput.bind("<Return>", getemail)
passwordInput.grid(row=2, column=1)

Label(root, text="It might take a while. Please wait...",
          bg="#fdfae5").grid(row=4, column=1)
Label(root, text="Processing 1,000 emails might take up to 7 minutes",
          bg="#fdfae5").grid(row=5, column=1)

Button(root, text="Download Excel Sheet", width=30, bg="#fcf4a3",
       command=getemail).grid(row=3, column=1)

root.mainloop()
