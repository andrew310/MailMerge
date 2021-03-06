

import wx
from Interface import MyFileDropTarget
from Merge import Merge
from Property import Property
from Emailer import Emailer
import os, sys

# begin wxGlade: dependencies
import gettext
# end wxGlade

# begin wxGlade: extracode
# end wxGlade


class MyFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MyFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        #self.text_ctrl_1 = wx.TextCtrl(self, wx.ID_ANY, "")
        self.button_1 = wx.Button(self, wx.ID_ANY, _("Start Letter Blast\n"))
        self.button_1.Bind(wx.EVT_BUTTON, self.onRadianButton)
        self.button_4 = wx.Button(self, wx.ID_ANY, _("Start Sales Blast"))
        self.button_2 = wx.Button(self, wx.ID_CLEAR, "")
        self.button_2.Bind(wx.EVT_BUTTON, self.onClearButton)

        #my insertions
        droptarget1 = MyFileDropTarget(self)
        self.tc_files = self.text_ctrl_1 = wx.TextCtrl(self, wx.ID_ANY, "", style=wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY)
        self.tc_files.SetDropTarget(droptarget1)
        self.dropped_files = []

        self.__set_properties()
        self.__do_layout()
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: MyFrame.__set_properties
        self.SetTitle(_("MailMonstrrr"))
        self.SetSize((380, 550))
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: MyFrame.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_1.Add(self.text_ctrl_1, 3, wx.EXPAND, 0)
        sizer_1.Add(self.button_1, 1, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_1.Add(self.button_4, 1, wx.EXPAND, 0)
        sizer_1.Add(self.button_2, 1, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade


    #my insertions 2:
    def onRadianButton(self, event):
        self.spreadsheet = 0
        #send spreadsheet path to Merge class
        myMerge = Merge(self.tc_files)
        #myMerge.hello()
        for i, v in enumerate(self.dropped_files):
            if self.dropped_files[i].endswith('.xlsx' or '.XLSX'):
                self.spreadsheet = v

        self.tc_files.WriteText("\n Processing Spreadsheet...")
        propertyList = myMerge.openFile(self.spreadsheet)
        #self.tc_files.WriteText("\n Creating PDFs...")
        myMerge.rename_files()
        self.tc_files.WriteText("\n Creating Emails...")
        myEmail = Emailer()
        myEmail.create_mail(propertyList)
        self.tc_files.WriteText("\n You're done! Don't forget to hit the clear button!")

    def onClearButton(self, event):
        self.tc_files.WriteText("\n Cleaning up files...")
        currentPath = os.path.dirname(sys.argv[0])
        folder = str(currentPath) + "/pdfs"

        for files in os.listdir(folder):
            file_path = folder + "/" + files
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except WindowsError:
                self.tc_files.WriteText("Could not remove file")

        self.tc_files.WriteText("DONE! Time to buy Andrew a coffee...")


    def notify(self, files, length, maxL):
        self.dropped_files.append(files)
        self.tc_files.WriteText(self.dropped_files[length])


    def notify2(self,files):
        self.tc_files.WriteText(files)


    def SetInsertionPointEnd(self):
        self.tc_files.SetInsertionPointEnd()

# end of class MyFrame
if __name__ == "__main__":
    gettext.install("app") # replace with the appropriate catalog name

    app = wx.App(0)
    #wx.InitAllImageHandlers()
    MailMonstrrr = MyFrame(None, wx.ID_ANY, "")
    app.SetTopWindow(MailMonstrrr)
    MailMonstrrr.Show()
    app.MainLoop()
