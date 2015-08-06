__author__ = 'Andrew.Brown'
from Mail import Mail
import win32com.client as win32
import os, sys
from Property import Property


class Emailer:

    def create_mail(self, mail_list):
        outlook = win32.Dispatch('Outlook.Application')
        pathName = os.path.dirname(sys.argv[0])
        print len(mail_list)
        for i in mail_list:
            myProp = i
            print myProp.propAddr

            #uses the template saved
            mail = outlook.CreateItemFromTemplate(pathName + r"/EmailTemplate.oft")
            #will insert agent name and address into email body
            mySpot = mail.GetInspector.WordEditor
            mySpot.Bookmarks("Name").Range.InsertAfter(myProp.agentName)
            mySpot.Bookmarks("Address").Range.InsertAfter(myProp.propAddr)

            #sets to/subject/attachment info
            mail.To = myProp.agentEmail
            mail.Subject = myProp.propAddr
            mail.Attachments.Add(myProp.attachmentPath + ".pdf")
            print "here's your attachmentPath"
            print myProp.attachmentPath
            #mail.HtmlBody = pathName + "/Email.docx"
            mail.Display(1)


    def create_blast(self, mail_list):
        Bookmarks("Date").Range.InsertAfter(date)
