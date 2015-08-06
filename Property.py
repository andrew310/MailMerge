__author__ = 'Andrew.Brown'


class Recipient:
    #default constructor parent class
    def __init__(self, name, email):
        self.agentName = name
        self.agentEmail = email

   #derived class for properties
class Property(Recipient):
     def __init__(self, fulladd, addr, city, state, zip, name, phone, email, attachment):
        Recipient.__init__(self, name, email)
        self.agentPhone = phone
        self.propAddr = fulladd
        self.address = addr
        self.city = city
        self.state = state
        self.zip = zip
        self.attachmentPath = attachment
