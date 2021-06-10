#!/usr/bin/env python

import win32com.client as client
import time


class MailSender:

    def __init__(self, receiver, ticket_number, attachment, worksheet):
        self.receiver = receiver
        self.ticket_number = ticket_number
        self.attachment = attachment
        self.worksheet = worksheet

    def mail_settings(self):

        if self.worksheet.upper() == "P":
            message_body = """
                            <html>
                              <head></head>
                              <body>
                                <font color="Black" size=-1 face="Arial">
                                <p>Dear SD Team,</p>
                                    <p>I&rsquo;m sending you an excel file in attachment. Please assist us:</p>
                                    <p>At the start of the migration the user has tagged these PST files whether the files are not belonging to him or are used by multiple users.</p>
                                        <ul>
                                            <li>For files with the status &ldquo;NOT MINE&rdquo; the user should provide the real owner of the file (in case he is aware of the owner), otherwise service 
                                            <br>desk should get in contact with the country IT for clarification.</li>
                                            <li>For files with the status &ldquo;SHARED&rdquo; the user should provide the users who are working together on the file (in case he is aware of the owners), 
                                            <br>otherwise service desk should get in contact with the country IT for clarification</li>
                                        </ul>
                                    <p><b>When check is performed, please send us updated excel file back. Do not create ticket as response! If something is unclear please, ask in mail response or contact someone from OMEGA team.</b></p>
                                    <p>Thank you</p>
                                </font>
                              </body>   
                            </html>
                       """
        elif self.worksheet.upper() == "I":
            message_body = """

                            <html>
                              <head></head>
                              <body>
                              <font color="Black" size=-1 face="Arial">
                                    <p>Dear SD Team,</p>
                                    <p>I am sending you the excel file &bdquo;IngestBlocked&ldquo; as an attachment. This file contains information about unsuccessful migration of PST files. Please do the following:</p>
                                        <ul>
                                            <li>The user has decided to migrate that file but the tool has not enough information to confirm that the file is belonging to the user, please double check whether that file is really belonging to the user, otherwise you should get in contact with the country IT for clarification</li>
                                        </ul>
                                    <p><b>When check is performed, please send us updated excel file back. Do not create ticket as response! If something is unclear please, ask in mail response or contact someone from OMEGA team.</b></p>
                                    <p>Thank you.</p>
                              </font>
                              </body>   
                            </html>
                            """
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
        message.CC = "itsupport@heidelbergcement.com"
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
#     mails = MailSender("xphrask01@vutbr.cz", "112233",
#                            "C:\\Users\\brani\\Downloads\\ParkedIngestBlockedFilesW49.xlsx")
