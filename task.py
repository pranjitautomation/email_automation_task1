import email
import imaplib
import json
import os
from datetime import datetime
from email.header import *

from imap_tools import MailBox


class EmailAutomation:
    
    def find_subject_lines(self, username, password, email_server):
        """_summary_

        Args:
            username (String): Use name of Email
            password (String): Password of the email

        Returns:
             A list: Having the subjects of all the companies
        """
        
        subject = []
        
        imap = imaplib.IMAP4_SSL(email_server)
        imap.login(username, password)
        imap.select('"[Gmail]/All Mail"', readonly = True) 
        
        response, messages = imap.search(None, 'UnSeen') #Selecting the unseen emails
        messages = messages[0].split()
        latest = int(messages[-1])

        for i in range(latest, latest-20, -1):
            res, msg = imap.fetch(str(i), "(RFC822)")
            
            for response in msg:
                if isinstance(response, tuple):
                    msg = email.message_from_bytes(response[1])
                if str(msg["Subject"]).count(":") == 2: #Taking the subject of only sales emails
                    subject.append(str(msg["Subject"]))
        
        return list(set(subject))
        
    def download_attachments_from_specific_subject(self, raw_subject,username,password):
        
        def create_dir_structure():
            #creating ddirectory according to company name
            if "attachments" not in os.listdir("./"):
                os.mkdir("attachments")
                print("Hello")

            #creating the folder name from subject of the mail
            raw_subject_split = raw_subject.split(":")
            parent_folder = raw_subject_split[-1].strip()
            sub_folder = raw_subject_split[0].strip()
            att_path = "attachments" + "\\" + parent_folder + "\\" + sub_folder
            
            if parent_folder not in os.listdir("./attachments"):
                os.makedirs(att_path)
            return att_path
        
        directory_path = create_dir_structure()
        
        def uid():
            now = datetime.now()
            print("now =", now)
            dt_string = now.strftime("%d%m%Y%H%M%S")
            return dt_string
    
        def date_time_object_2_raw(date_time_object):
            # making File name according to date of the email
            return str(date_time_object.strftime("%d%b%Y")) + "-Sales" + "_" + str(uid()) + ".xlsx"
        
        def download_and_delete():
        # Downloading the excel present in sales report mail & delete the seen mail
            
            with MailBox('imap.gmail.com').login(username, password, 'INBOX') as mailbox:
                for msg in mailbox.fetch():
                    
                    if str(msg.subject) == raw_subject:
                        date_of_msg = msg.date

                        for att in msg.attachments:
                            # print(att.filename, att.content_type)
                            filename = date_time_object_2_raw(date_of_msg)
                            with open(directory_path + "\\" + filename, 'wb') as f:
                                f.write(att.payload)
                        mailbox.delete(msg.uid)

        download_and_delete()

#taking vault data from json file
path_to_json = "./vault.json"
with open(path_to_json, "r") as handler:
    info = json.load(handler)

username = info["username"]
password = info["password"]
email_server = "imap.gmail.com"


eobj = EmailAutomation()
raw_subject = eobj.find_subject_lines(username, password, email_server)

for x in raw_subject:
    eobj.download_attachments_from_specific_subject(x,username,password)