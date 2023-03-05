import smtplib
import speech_recognition as sr
import pyttsx3
import imaplib
import os
# from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import email
import streamlit as st

attach_folder="C:/Users/HP/Desktop/"

listener = sr.Recognizer()
engine = pyttsx3.init()

EMAIL_ID = '19z210@psgtech.ac.in'  
PASSWORD = 'yigbwqyubutpwbue'

def talk(text):
    engine.say(text)
    engine.runAndWait()

st.markdown("<h1 style='text-align: center; color: skyblue;'>InterComm</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: white;'>Voice based Email Assistant</h4><br>", unsafe_allow_html=True)
talk("Welcome to ... Voice based Email Assistant  ")


col1, col2, col3, col4, col5 = st.columns([1,1,1,1,1])
with col1:
    button1=st.button('Compose')
with col2:
    button2=st.button('Inbox')
with col3:
    button3=st.button('Spam')
with col4:
    button4=st.button('Trash')
with col5:
    button5=st.button('Exit')



def listen():
    try:
        with sr.Microphone() as source:
            print('listening...')
            talk("Speak...")
            voice = listener.listen(source)
            info = listener.recognize_google(voice)
            print(info)
            return info.lower()
    except:
        pass

def spechar(text):
    temp=text
    special_chars = ['attherate','dot','underscore','space']
    for character in special_chars:
        while(True):
            pos=temp.find(character)
            if pos == -1:
                break
            else :
                if character == 'attherate':
                    temp=temp.replace(' attherate ','@')
                elif character == 'dot':
                    temp=temp.replace('dot','.')
                elif character == 'underscore':
                    temp=temp.replace(' underscore ','_')
                elif character == 'space':
                    temp=temp.replace(' space ',' ')
    return temp

def read_files_with_name(name):
    directory= "C:/Users/HP/Desktop/"
    for root, dirs, files in os.walk(directory):
        for file in files:
            if name in file:
                print(os.path.join(root, file))
                talk(file)
    talk("Name the file you want to send as an attachment.")
    fname=listen()
    filefinal=spechar(fname)
    return filefinal

def send_email(receiver, subject, message):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(EMAIL_ID,PASSWORD )
    mail = MIMEMultipart()
    mail['From'] = '19z210@psgtech.ac.in'
    mail['To'] = receiver
    mail['Subject'] = subject
    mail.attach(MIMEText(message, 'plain'))
    path = "C:/Users/HP/Desktop/"
    talk("Do you want to send an attachment?")
    ans=listen()
    if ans=="yes":
        talk("Say file name..")
        ftemp=listen()
        ftemp=spechar(ftemp)
        talk("The following files with the filename name you mentioned have been found in the directory. Choose the file you want to send by naming the exact file with extension.")
        filefinal = read_files_with_name(ftemp)
        filename= os.path.join(path, filefinal)

        # Attach the file
        basename = os.path.basename(filename)

        attachment = open(filename, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % basename)
        mail.attach(part)
        server.send_message(mail)
    else:
        server.send_message(mail)

email_list = {
    'deepti':'deeptiravikumar@gmail.com',
    'shruti' : 'shrutiravikr@gmail.com'
}


def draft():
    talk('To Whom you want to send email')
    name = listen()
    receiver = email_list[name]
    print(receiver)
    talk('What is the subject of your email?')
    subject = listen()
    talk('Tell me the text in your email')
    message = listen()
    send_email(receiver, subject, message)
    talk('Your email has been sent succesfully')
    talk('Do you want to send more email?')
    send_more = listen()
    if 'yes' in send_more:
        draft()
    else: 
        talk("Redirecting to the functions menu")
        main()

def download_attachments(M, msg):
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        att_path = os.path.join(attach_folder, filename)

        if not os.path.isfile(att_path):
            fp = open(att_path, 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
        talk("The email has an attachment")
        talk(f"Downloaded attachment {filename} "+"to desktop")


def get_unread_mail_count(M):
    stat, total = M.select("inbox")
    stat, unread_emails = M.search(None, "UNSEEN")
    unread_count = len(unread_emails[0].split())
    return unread_count

def get_unread_mail_contents(M):
    stat, total = M.select("inbox")
    stat, unread_emails = M.search(None, "UNSEEN")
    unread_mail_ids = unread_emails[0].split()
    for email_id in unread_mail_ids:
        stat, email_data = M.fetch(email_id, "(RFC822)")
        for response in email_data:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                for header in ["From", "Subject", "Date"]:
                    talk(f"{header}: {msg[header]}")
                download_attachments(M,msg)
    talk("Redirecting to functions menu")
    main()
    

def check_inbox():
    M = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    M.login(EMAIL_ID, PASSWORD)

    unread_count = get_unread_mail_count(M)
    if unread_count > 0:
        talk(f"You have {unread_count} , unread emails.")
        talk("Do you want to read the unread emails?")
        ans=listen()
        if ans=="yes":
            get_unread_mail_contents(M)
    else:
        talk("You have no unread emails. Redirecting to functions menu")
        main()

    talk("Redirecting to functions menu")
    main()

    M.close()
    M.logout()

def get_mail_ids(M, mailbox):
    stat, data = M.select(mailbox)
    stat, mail_ids = M.search(None, "ALL")
    return mail_ids[0].split()

def read_mail_headers(M, mail_ids, mailbox):
    for email_id in mail_ids:
        stat, email_data = M.fetch(email_id, "(RFC822)")
        for response in email_data:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                talk(f"From: {msg['From']}\nSubject: {msg['Subject']}")

def delete_emails(M, mail_ids):
    for email_id in mail_ids:
        M.store(email_id, "+FLAGS", "\\Deleted")
    M.expunge()

def check_delete_emails(M, mail_ids, mailbox):
    talk(f"Do you want to delete the emails in {mailbox}? (yes or no)")
    delete_confirm = listen()
    if delete_confirm == "yes":
        talk(f"Deleting mails in {mailbox}. Please Wait")
        delete_emails(M, mail_ids)
        talk(f"Emails in {mailbox} have been deleted.")
    else:
        talk(f"Emails in {mailbox} have not been deleted.")

def check_trash():
    M = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    M.login(EMAIL_ID, PASSWORD)

    mail_ids = get_mail_ids(M, "[Gmail]/Trash")
    if mail_ids:
        talk("The following are the emails in the Trash folder:")
        read_mail_headers(M, mail_ids, "Trash")
        check_delete_emails(M, mail_ids, "Trash")
    else:
        talk("Trash folder is empty.")

    M.close()
    M.logout()

    talk("Redirecting to functions menu")
    main()

def check_spam():
    M = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    M.login(EMAIL_ID, PASSWORD)
    mail_ids = get_mail_ids(M, "[Gmail]/Spam")
    if mail_ids:
        talk("The following are the emails in the Spam folder:")
        read_mail_headers(M, mail_ids, "Spam")
        check_delete_emails(M, mail_ids, "Spam")
    else:
        talk("Spam folder is empty.")

    M.close()
    M.logout()

    talk("Redirecting to functions menu")
    main()


def main():

    talk("What would you like to do ? To compose an email say compose. To open Inbox folder say Inbox. To open Spam folder say Spam. To open Trash folder say Trash. To exit mail assistant say exit.")
    choice = listen()
    if choice == 'compose' or choice=='pose':
        draft()
    elif choice== 'inbox' or choice=='box':
        check_inbox()
    elif choice == "spam" or choice == "span":
        check_spam()
    elif choice == "trash" or choice=='rash':
        check_trash()
    elif choice=="exit":
        talk("Thank you for using the service. Redirecting the to main menu")
        exit()
    else:
        talk("Wrong choice. Please say only the option")
        main()


if __name__ == '__main__':
    main()
