__author__ = 'Andrew.Brown'
import win32com.client as win32

class Mail:
    def __init__(self, text, subject, attachment, name, email):
        self.text = text
        self.subject = subject
        self.attachment = attachment
        self.name = name
        self.email = email



