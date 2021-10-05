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

        if self.worksheet.upper() == "FS":
            message_body = """
                <p>Dear SD Team,</p>
                <p>I&rsquo;m sending you an excel file in the attachment. Please assist us:</p>
                <p>Please decide on the files on the file server whether the file can be excluded from the PST migration (please delete the file from its location) or whether the file shall be migrated to a user or shared mailbox (if the shared mailbox don&rsquo;t exist, please request the shared mailbox)</p>
                <p>In case of any technical questions regarding the migrations please, contact <strong>Ettling, Andreas (Heidelberg) DEU.</strong></p>
                <p><strong>When the check is performed, please send updated excel file to <span style="color: red;"> Andreas Ettling (andreas.ettling@heidelbergcement.com).</span> Do not create ticket in response!</strong></p>
                <p>Thank you</p>
                 """
        elif self.worksheet.upper() == "A":
            message_body = """
                <p>Dear SD Team,</p>
                <p>I&rsquo;m sending you an excel file in the attachment. Please assist us:</p>
                <p>There is no feedback received from the user. This can be caused by issues between the agent on the client and the migration server. Please restart the agent on the client or restart the client itself or ask the user how to proceed with the PST files detected for him and provide us the feedback and we initiate the migration of the files from the server.</p>
                <p>To restart the agent please run on the client &ldquo;C:\Program Files (x86)\Quadrotech\Migration Agent\ ResetMigrationAgent.exe&rdquo;</p>
                <p>In case of any technical questions regarding the migrations please, contact <strong>Ettling, Andreas (Heidelberg) DEU.</strong></p>
                <p><strong>When the check is performed, please send updated excel file to <span style="color: red;"> Andreas Ettling (andreas.ettling@heidelbergcement.com).</span> Do not create ticket in response!</strong></p>
                <p>Thank you</p>
                <p>&nbsp;</p>
                """

        elif self.worksheet.upper() == "O":
            message_body = """
                <p>Dear SD Team,</p>
                <p>I&rsquo;m sending you an excel file in the attachment. Please assist us:</p>
                <ul>
                <li>Column status &ldquo;Owner unclear&rdquo;: the tool has not enough information to be sure on the owner &ndash; please contact the users and ask whether they are the owner of the file or the file can be excluded from the migration</li>
                <li>Column status &ldquo;User is not owner&rdquo;: ask the user about the real owner of the file or whether the file can be excluded from the migration</li>
                <li>Column status &ldquo;Shared PST&rdquo;: ask the user who shall work with the PST file, check whether a shared mailboxes with proper permissions is existing or request a new mailbox and provide us the name of the shared mailbox or whether the file can be excluded from the migration</li>
                </ul>
                <p>In case of any technical questions regarding the migrations please, contact <strong>Ettling, Andreas (Heidelberg) DEU.</strong></p>
                <p><strong>When the check is performed, please send updated excel file to <span style="color: red;"> Andreas Ettling (andreas.ettling@heidelbergcement.com).</span> Do not create ticket in response!</strong></p>
                <p>Thank you</p>
                """

        elif self.worksheet.upper() == "N":
            message_body = """
                <p>Dear SD Team,</p>
                <p>I&rsquo;m sending you an excel file in the attachment. Please assist us:</p>
                <p>These users are not yet enabled for the PST migration.</p>
                <ul>
                <li>Users marked in red are no longer existing, please reply whether the files can be excluded from the PST migration or have to be assigned to a different user</li>
                <li>Users marked in orange are in the terminated users OU, please reply whether the files can be excluded from the PST migration or have to be assigned to a different user</li>
                <li>Users marked in green are regarded as completed &ndash; the system hasn&rsquo;t seen the PST files detected for that users for more than 30 days. The files might have been deleted by the user or there is an issue with the agent that the agent is not able to report to the server. Please restart the agent on the client in that case. Please request a confirmation from the user that the PST files are no longer needed or have to be migrated.</li>
                </ul>
                <p>In case of any technical questions regarding the migrations please, contact <strong>Ettling, Andreas (Heidelberg) DEU.</strong></p>
                <p><strong>When the check is performed, please send updated excel file to <span style="color: red;"> Andreas Ettling (andreas.ettling@heidelbergcement.com).</span> Do not create ticket in response!</strong></p>
                <p>Thank you</p>
                <p>&nbsp;</p>
            """

        elif self.worksheet.upper() == "FI":
            message_body = """
                <p>p&gt;Dear Team,</p>
                <p>I&rsquo;m sending you an excel file in the attachment. Please assist us:</p>
                <p>Not all mails of the PST files could be ingested into the online archive. The sheet is showing the amount of items in the PST file and how many items couldn&rsquo;t be ingested. Please verify with the user whether we can skip the items not migrated.</p>
                <p>Please verify with the user whether we can skip the items not migrated. The user can verify the details on the items not ingested by himself:</p>
                <ul>
                <li>Self service - QUADROtech (heidelbergcement.com)</li>
                <li>In the column &ldquo;Progress&rdquo; the symbol &ldquo;!&rdquo; indicates files with failed items</li>
                <li>Click on the symbol in the column &ldquo;Info&rdquo; for this file</li>
                <li>A popup window opens and at the tab &ldquo;Failed Items Details&rdquo; are all failed items listed</li>
                </ul>
                <p>In case there is a very high amount of items which can&rsquo;t be ingested it might be caused by the fact that the Online Archive of the user reached it&rsquo;s limit and after Microsoft has extended the limit of the Online Archive the PST migration tool might be able to migrate further items. In the comment column we added a not which kind of items are missing.</p>
                <p>In case of any technical questions regarding the migrations please, contact <strong>Ettling, Andreas (Heidelberg) DEU.</strong></p>
                <p><strong>When the check is performed, please send updated excel file to <span style="color: red;"> Andreas Ettling (andreas.ettling@heidelbergcement.com).</span> Do not create ticket in response!</strong></p>
                <p>Thank you</p>
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
        #message.CC = "itsupport@heidelbergcement.com"
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
