__author__ = 'Andrew.Brown'

import wx

class MyFileDropTarget(wx.FileDropTarget):
    """"""
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.count = 0

    def OnDropFiles(self, x, y, filenames):
        self.window.SetInsertionPointEnd()
        self.window.notify2("\n%d file(s) dropped at %d,%d:\n" %
                              (len(filenames), x, y))
        for file in filenames:
            self.window.notify(file, self.count, len(filenames))
            self.count = self.count+1



