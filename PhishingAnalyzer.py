# Colby Austin
# This script is to help automate the analysis and response process in phishing analysis.
#
# v. 0.1
#   The script will loop through all the emails in the analyst phishing inbox, variables are statically set
#
# v. 0.2
#   The script now parses the email to help the analyst identify the email in question
#
# v 0.3
#   The script works properly and can now send the appropriate response after the analyst classifies the email
#
# v 0.4
#   implemented config file... broke everything
#
# v 0.5
#   found missing reference
#
# v 0.6
#   implemented signature file, to make script easier to share
#
#v 1.0
#   shared with everyone
#
#
# Future things to add:
#       grammar checks for the email
#       create a local db to store past malicious addresses and keep stats on them
#       Create log file to track what the analyst response is to the emails
#       create auto response for emails that match phishing emails that have already been analyzed
#       learn api for urlscan.io to see if you can automate the process of scanning urls
#
#Long term goal:
#       Create GUI to display all of the data in an easy to read manner
########################################################################################################################

import win32com.client
import os
from spellchecker import SpellChecker
from datetime import datetime


def add_log(Sender,Subject,Response):
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    log.write("\n\n#############\n{0}".format(dt_string))
    log.write("Sender:          {0}".format(Sender))
    log.write("Subject:         {0}".format(Subject))
    log.write("Classified as:   {0}".format(Response))


def scan_email(message):         #signs to check for in the email that could point to phishing
    wordlist = message.body.split()             #begin spell check
    spell = SpellChecker()
    amount_miss = len(list(spell.unknown(wordlist)))
    print("Possible misspelled words/total number of words: {0}/{1}" .format(amount_miss, len(wordlist)))       #print spell check results
    #print("Sender address: {0}    Return address: {1}".format(message.sender, message.) )                #check that return address is the same as sending address


def classify(message):
    recipient = message.Sender
    for header_line in message.body.split("\n"):
        if "From: " in header_line:
            sender = header_line.split(": ")[1]
            sender = sender.rstrip()
        if "Subject: " in header_line:
            subject = header_line.split(": ")[1]
            subject = subject.rstrip()
    print('Recipient: {}'.format(recipient))
    try:
        print('Sender: {}'.format(sender))
    except UnboundLocalError:
        print("ERROR:\nSomething is funky with the email header")
        print("Email will not send properly")
    try:
        print('Subject: {}'.format(subject))
    except UnboundLocalError:
        print("ERROR:\nSomething is funky with the email header")
        print("Email will not send properly")
    scan_email(message)
    print()
    print("Analyze the results and then Choose the classification")
    print("1: Malicious")
    print("2: Legit")
    print("3: Not Malicious, Possibly Legit")
    print("4: Spam")
    print("5: Suspicious")
    print("6: Skip ")
    Classification = input("Enter the number corresponding to your classification: ")
    reply = message.ReplyAll()
    reply._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))                  #used to send from a specific email address when you have multiple shared addresses on one outlook account
    sender = sender.replace("\n", "")
    subject = subject.replace("\n", "")
    ########################################################################################################################
    if Classification == "1":  # malicious
        reply.Body = """Good afternoon,
Information Security has determined an email you received from [{0}] with the subject [{1}] is MALICIOUS.
If you replied to the sender, opened any link or attachment associated with the email, then please respond to this email so we can investigate further.
\n{2}""".format(sender, subject, signature)
        if verify == "0":
            reply.Send()
        else:
            reply.Display()
        add_log(sender, subject, "Malicious")
    ########################################################################################################################
    elif Classification == "2":  # Legit
        reply.Body = """Good afternoon,
Information Security has determined an email you received from [{0}] with the subject [{1}] is LEGITIMATE. 
Your submission helps improve our filtering tools and reduce further unwanted emails.
We appreciate your assistance in keeping NRF safe and secure!
\n{2}""".format(sender, subject, signature)
        if verify == "0":
            reply.Send()
        else:
            reply.Display()
        add_log(sender, subject, "Legit")
########################################################################################################################
    elif Classification == "3":  # Not Malicious, Possibly Legit
        reply.Body = """Good afternoon,
Information Security has determined an email you received from [{0}] with the subject [{1}] does not contain any malicious attachments or links, however we are unable to determine the legitimacy of the email.
Some questions to help decide the legitimacy of common emails are listed below to assist you.
    1. Did the email come from a person/company that you work with?
    2. Were you expecting an email from this group with an attachment/link?
    3. Is this a service/application that you use?
If you are still unsure of the email you can reach out to the person that sent it and ask them directly.
You can recover the original email from your Deleted Items folder in Outlook.  
Please use caution if you do respond, and forward any future emails to phishing@nortonrosefulbright.com to review before opening if you are unsure. 
We appreciate your assistance in keeping NRF safe and secure!
\n{2}""".format(sender, subject, signature)
        reply.Display()
        add_log(sender, subject, "Not Malicious, Possibly Legit")
########################################################################################################################
    elif Classification == "4":  # Spam
        reply.Body = """Good Afternoon,
Information Security has determined an email you received from [{0}] with the subject [{1}] is SPAM. 
Please follow these instructions (Phishing Information) on Athena if you wish to block the sender from sending you further emails.
Your submission helps improve our filtering tools and reduce further unwanted emails.
We appreciate your assistance in keeping NRF safe and secure!
\n{2}""".format(sender, subject, signature)
        if verify == "0":
            reply.Send()
        else:
            reply.Display()
        add_log(sender, subject, "Spam")
########################################################################################################################
    elif Classification == "5":  # Suspicious
        reply.Body = """Good afternoon,
Information Security has determined an email you received from [{0}] with the subject [{1}] is SUSPICIOUS and should be handled with caution.
If you replied to the sender, opened any link or attachment associated with the email, please reply to this email so we can investigate further.
\n{2}""".format(sender, subject, signature)
        if verify == "0":
            reply.Send()
        else:
            reply.Display()
        add_log(sender, subject, "Suspicious")
########################################################################################################################
    elif Classification == "6":  # Skip
        print("Moving on then.")
        add_log(sender, subject, "skip")


if __name__ == '__main__':
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook2 = win32com.client.Dispatch("Outlook.Application")
    log = open('analyzed_emails.log', 'a+')
    log.write("########################################################################################################################")
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    log.write(dt_string)
    log.write("Begin")
    for line in open('config.conf', 'r'):
        if "Email_Address: " in line:
            email_address = line.split(": ")[1].strip("\n")
        if "Email_Folder: " in line:
            folder = line.split(": ")[1].strip("\n")
        if "Verify_email: " in line:
            verify = line.split(": ")[1].strip("\n")
    phishing_reports = outlook.Folders[''].Folders[''].Folders[''].Folders[folder].Items        #this is the file path
    send_account = None
    for account in outlook2.Session.Accounts:
        if account.DisplayName == email_address:
            send_account = account
            break
    with open('signature.txt', 'r') as file:
        signature = file.read()
    for phish in phishing_reports:
        classify(phish)
        print("##################################################################################################################################")
