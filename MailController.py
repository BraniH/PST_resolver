#!/usr/bin/env python

import win32com.client as client
import time
import os


def open_file():
    all_messages = []
    with open(os.getcwd() + "\\Mail_texts.txt") as messages:
        for message_line in messages:
            all_messages.append(message_line)

        return all_messages


def get_text(choice):
    all_messages = open_file()

    message_body = ""
    start = 0
    counter = 0
    message_status = False
    for message in all_messages:
        if choice in message:
            start = counter
            message_status = True

        if not message_status:
            pass
        elif (counter > start) and ("!!!!!!!END!!!!!!!" == message.replace("\n", "")):
            message_status = False
        elif (counter > start) and ("!!!!!!!END!!!!!!!" != message):

            message_body += message

        counter += 1

    return message_body


class MailSender:

    def __init__(self, receiver, ticket_number, attachment, worksheet):
        self.receiver = receiver
        self.ticket_number = ticket_number
        self.attachment = attachment
        self.worksheet = worksheet

    def mail_settings(self):
        if self.worksheet.upper() == "FS":
            message_body = get_text("----File Server----")

        elif self.worksheet.upper() == "A":
            message_body = get_text("----Awaiting - Discovered----")

        elif self.worksheet.upper() == "O":
            message_body = get_text("----Owner unclear - User disagree----")

        elif self.worksheet.upper() == "N":
            message_body = get_text("----Not Enabled Users----")

        elif self.worksheet.upper() == "FI":
            message_body = get_text("----Failed Items----")
        else:
            print(" ¯\_(ツ)_/¯ Something went wrong, closing!")
            time.sleep(3)
            exit()

        mail_subject = "PST migration: " + str(self.ticket_number)

        return mail_subject, message_body

    def sent_mail(self):
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.Display()
        message.To = self.receiver
        # message.CC = "itsupport@heidelbergcement.com"
        message.SentOnBehalfOfName = "itsupport@heidelbergcement.com"
        subject, body = self.mail_settings()
        message.subject = subject
        sign = message.HTMLbody
        message.HTMLbody = body + sign
        message.Attachments.Add(self.attachment)
        message.Save()
        message.Send()
        return self.attachment.split("\\")[-1] + " ---> sent to " + self.receiver


# if __name__ == '__main__':
