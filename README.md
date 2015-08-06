# MailMerge
An email blast / mail merge program written in python using pywin32 lib. Uses excel, outlook, and word.

This is a standalone Windows app -- I have included a py2exe script which builds the program into an executable. 

The tool is geared towards realtors, as you can see from the MOCK_DATA sheet it contains a list of addresses and names. 
The address, name and date are added to the letter, which is converted to PDF and then added as an attachment to the email.

The email template can be edited, as can the attachment template. The whole program should be easy to tweak for a number of uses 
and projects. 

wxPython was used for a simple drag and drop interface. The "Sales Blast" button doesn't do anything yet, but the top button 
will start the letter blast.
