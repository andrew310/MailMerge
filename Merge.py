__author__ = 'Andrew.Brown'
import shutil
import urllib
import sys, os
import datetime
import win32com.client as win32
from win32com import client
from win32com.client import constants

#this is my own defined class for property objects
from Property import Property
from time import strftime
import wx


class Merge:

    def __init__(self, text_control):
        self.tc = text_control

    #def hello(self):
        #print "hello from the merge class"
        #return

    #simply opens the file and passes it to getData, setData
    def openFile(self, spreadsheet):
        # open instance of excel
        win32.gencache.EnsureModule('00020813-0000-0000-C000-000000000046', 0, 1, 8)
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = 0  # don't want to actually open the program
        print "excel and word opened successfully"
        workBook = excel.Workbooks.Open(spreadsheet)
        workSheet = workBook.Worksheets("Sheet1")

        #pList will contain the array returned from getData
        pList = Merge.getData(self, workSheet)
        # close out of excel
        excel.Application.Quit()

        word = win32.Dispatch('Word.Application')

        #makes a copy of the template so as not to destroy the original
        for i in range(0, len(pList)):
            myProp1 = pList[i]
            print myProp1.propAddr
            curPath = os.path.dirname(sys.argv[0])
            src = str(curPath) + '\\Template.docx'
            dest = str(curPath) + '/tmpw/Template.docx'
            shutil.copyfile(src, dest)
            doc = word.Documents.Open(dest)
            #assert isinstance(doc, object) look at this later
            #call the set data class
            Merge.setData(self, doc, myProp1)
            i+=1
        #return list of properties

        word.Application.Quit()
        return pList
    #end of openFile function

    #takes worksheet as argument, counts the rows, gets data
    def getData(self, ws):
        myWS = ws
        iCount = 2  # used to set the range in iRange, set to 2 because of header
        # boolean to stop the loop
        iStop = False;
        #start an empty array for the properties
        propList = []

        while iStop == False:

            #check to see if this is the last loop
            endCheck = myWS.Range("A" + str(iCount+1) + ":" + "A" + str(iCount+1))
            checkVal = endCheck.Value
            if (checkVal == None):
                iStop = True

            # get the address
            iRange = myWS.Range("A" + str(iCount) + ":" + "A" + str(iCount))
            values = iRange.Value

            # get the address
            iRange = myWS.Range("A" + str(iCount) + ":" + "A" + str(iCount))
            values = iRange.Value
            # get the city
            cRange = myWS.Range("B" + str(iCount) + ":" + "B" + str(iCount))
            city = cRange.Value
            # get the state
            sRange = myWS.Range("C" + str(iCount) + ":" + "C" + str(iCount))
            state = sRange.Value
            # get the zip
            zRange = myWS.Range("D" + str(iCount) + ":" + "D" + str(iCount))
            zip = zRange.Value
            # the zip has a .0 on the end, so I deleted that
            stringZip = str(zip)
            stringZip = stringZip[:-2]

            # this holds the entire address
            listAddress = [str(values), ", ", str(city), ", ", str(state), " ", stringZip]
            iAddress = "".join(listAddress)

            # get the agent name
            nRange = myWS.Range("E" + str(iCount) + ":" + "E" + str(iCount))
            name = nRange.Value

            # get the agent phone
            pRange = myWS.Range("F" + str(iCount) + ":" + "F" + str(iCount))
            phone = pRange.Value

            # get the agent email
            eRange = myWS.Range("G" + str(iCount) + ":" + "G" + str(iCount))
            email = eRange.Value

            #names the file and saves path to attachmentPath
            pathName = os.path.dirname(sys.argv[0])
            savePath = [str(pathName), "\\pdfs\\", str(iAddress)]
            #note: does not contain ".docx" or ".pdf"
            totalPath = "\\".join(savePath)

            #create a property object for each row
            myProp = Property(str(iAddress), str(iRange), str(city), str(state), str(zip), str(name), str(phone), str(email), str(totalPath))
            #add property to the list
            propList.append(myProp)

            #increment counters
            iCount += 1
        # end of loop

        #returns the array to calling function
        return propList
        # end of the getData function

        #################  setData function ###################
        #inserts data from the property object list into separate word documents
        #takes a word doc template, and a list of Property objects
    def setData(self, docx, xProp):
        #get the current date
        today = datetime.date.today()
        uDate = [str(today.month), str(today.day), str(today.year)]
        #join together the dates with forward slashes
        date = "/".join(uDate)
        #insert date into word doc
        docx.Bookmarks("Date").Range.InsertAfter(date)


        formatAddress = ["\n", xProp.address, "\n", xProp.city, ", ", xProp.state, " ", xProp.zip[:-2]]
        formatAddress = "".join(formatAddress)
        #insert the address into the word doc
        docx.Bookmarks("Address").Range.InsertAfter(formatAddress)

        #insert the agent name into the document
        docx.Bookmarks("Name").Range.InsertAfter(xProp.agentName)

        #names the file and saves it
        totalPath = xProp.attachmentPath + ".docx"
        #sets the attachmentPath
        docx.SaveAs(totalPath)
        docx.Close()

        return xProp

    # #function that takes filetype and folder path, will count files in that dir
    # def count_files(self, filetype, folder):
    #     count_files = 0
    #     for files in os.listdir(folder):
    #         if files.endswith(filetype):
    #             count_files += 1
    #     return count_files

    #converts from word to pdf, all the word documents in a folder
    def rename_files(self):
        #prints message to text control box
        self.tc.WriteText("\n Creating PDFs")
        currentPath = os.path.dirname(sys.argv[0])
        folder = str(currentPath) + "\pdfs"
        word = client.DispatchEx("Word.Application")
        for files in os.listdir(folder):
            if files.endswith(".docx"):
                new_name = files.replace(".docx", r".pdf")
                in_file = os.path.abspath(folder + "\\" + files)
                new_file = os.path.abspath(folder + "\\" + new_name)
                doc = word.Documents.Open(in_file)
                outMSG = strftime("%H:%M:%S"), " docx -> pdf " + os.path.relpath(new_file)
                print outMSG
                self.tc.WriteText('\n' + str(outMSG))
                doc.SaveAs(new_file, FileFormat = 17)
                doc.Close()
            if files.endswith(".doc"):
                new_name = files.replace(".doc", r".pdf")
                in_file = os.path.abspath(folder + "\\" + files)
                new_file = os.path.abspath(folder + "\\" + new_name)
                doc = word.Documents.Open(in_file)
                print strftime("%H:%M:%S"), " doc  -> pdf ", os.path.relpath(new_file)
                doc.SaveAs(new_file, FileFormat = 17)
                doc.Close()



