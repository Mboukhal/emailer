import sys
import os
import subprocess
import signal
import email, imaplib
import base64
from email.mime import multipart
from tkinter import *
from functools import partial
import pystray
import PIL.Image
import time
from threading import Thread

IMAP_SERVER = 'imap.gmail.com'
IMAP_PORT = 993
EMAIL_FOLDER = 'INBOX'
OUTPUT_DIRECTORY = 'save_mail'

image = PIL.Image.open('R.png')

def open_save(icon, item):
    subprocess.call(f'explorer {OUTPUT_DIRECTORY}"', shell=True)


def on_clicked(icon, item):
    os.kill(os.getpid(), signal.SIGINT)

icon = pystray.Icon("Emailer", image, menu=pystray.Menu(
    pystray.MenuItem("Open seved folder", open_save),
    pystray.MenuItem("Exit", on_clicked)
    ))

flag = 1
try:
    f = open('driver.exe', 'r')
    if len(f.read()) == 0:
        flag = 0
    f.close()
except:
    flag = 0

if flag == 0:
    tkwin = Tk()
    status = Label(tkwin, text='').grid(row=4, column=2)

open_url = 'https://support.google.com/accounts/answer/185833'


def x(p, u):
    global tkwin
    global status
    defaultbg = tkwin.cget('bg')
    if len(str(u.get())) < 10 or len(str(p.get())) < 8:
        status = Label(tkwin, text='Error', fg='red', bg=defaultbg).grid(row=4, column=2)
        return
    if str(u.get())[-10:] != '@gmail.com':
        status = Label(tkwin, text='Error', fg='red', bg=defaultbg).grid(row=4, column=2)
        return

    M = imaplib.IMAP4_SSL(host=IMAP_SERVER, port=IMAP_PORT)
    try:
        M.login(str(u.get()), str(p.get()))
        x0 = 1
    except Exception as ms:
        if str(ms).split(' ')[0] == "b'[ALERT]":
            os.system(f"start /MIN {open_url}")
            status = Label(tkwin, text='Activation', fg='purple', bg=defaultbg).grid(row=4, column=2)
            x0 = 0
        elif str(ms).split(' ')[1] == 'Invalid':
            status = Label(tkwin, text='Error', fg='red', bg=defaultbg).grid(row=4, column=2)
            x0 = 0
    if x0 == 1:
        f = open("driver.exe", mode='w+', encoding='utf-8')
        f.write(str(u.get()))
        f.write('\n')
        f.write(str(p.get()))
        f.close()
        tkwin.destroy()
        start(u.get(), p.get())
    M.logout()
    

def login_win():
    
    global tkwin
    global status

    tkwin.geometry('250x80')
    tkwin.title('Emailer')

    user = Label(tkwin, text='Gmail: ').grid(row=0, column=0)
    username = StringVar()
    userEntry = Entry(tkwin, textvariable=username).grid(row=0, column=1)

    password = Label(tkwin, text='Password: ').grid(row=1, column=0)
    password = StringVar()
    passEntry = Entry(tkwin, textvariable=password, show='-').grid(row=1, column=1)
    status = Label(tkwin, text='').grid(row=4, column=2)

    x1 = partial(x, password, username)

    loginb = Button(tkwin, text='Start', command=x1).grid(row=4, column=1)

    tkwin.mainloop()
    
    return password, username

def get_email(msg):
    source = msg['from']
    to = msg['to']
    subject = msg['subject']
    att = msg['attach']
    count = 1
    for part in msg.walk():
        if part.get_content_maintype() == multipart:
            continue
        filename = part.get_filename()
        if not filename:
            ext = '.html'
            filename = 'msg-part%80d%s' %(count, ext)
        count += 1
    ct = part.get_content_type()

def remove_all(M):

    M.search(None, "ALL")

def process_mailbox(M):

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        return
    for num in data[0].split():
        num1 = num
        rv, data = M.fetch(num, '(RFC822)')
        eml = data[0][1].decode('utf-8')
        email_msg = email.message_from_string(eml)
        
        x = email_msg['date'].split(' ')
        x1 = x[1]+ " "+ x[2] + " " +x[3] + " " + email_msg['from'].split(' ')[2][1:-1]
        path = os.path.join(OUTPUT_DIRECTORY, x1)
        try:
            os.mkdir(path, mode=0o777)
        except:
            pass
        num = email_msg['from'].split(' ')[2][1:-1]
        if rv != 'OK':
            return
        file_name = '%s\\%s' %(path, num)
        f = open(file_name + '.eml' , 'wb')
        f.write(data[0][1])
        f.close()
###########################################################
        # save attachment
###########################################################
        
        for part in email_msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            d_file = part.get_filename()
            if bool(d_file):
                pfile = os.path.join(path, d_file)
                if not os.path.isfile(pfile):
                    fp = open(pfile, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
###########################################################
        subprocess.call(f'ecc.exe "{file_name}.eml"', shell=True)
        subprocess.call(f'del "{file_name}.eml"', shell=True)
        #M.store(num1, '+FLAGS', '\\Deleted')
        #M.expunge()


def main(u, p):
    while 1:
        M = imaplib.IMAP4_SSL(host=IMAP_SERVER, port=IMAP_PORT)
        try:
            M.login(u, p)
        except Exception:
            os.kill(os.getpid(), signal.SIGINT)
        rv, data = M.select(EMAIL_FOLDER)
        if rv == 'OK':
            process_mailbox(M)
            M.close()
        M.logout()
        time.sleep(1800*60)
        
def start(us, pa):
    t1 = Thread(target=icon.run)
    t1.start()
    main(us, pa)

if __name__ == "__main__":
    
    if flag == 0:
        b, a = login_win()
        pa = a.get()
        us = b.get()
    else:
        f = open('driver.exe', mode='r', encoding='utf-8')
        us = str(f.readline())
        pa = str(f.readline())
        f.close()
    start(us, pa)
