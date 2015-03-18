#coding=utf8

__author__ = 'FuQiang'

import os
import os.path
import opExcel

import wx
import  wx.lib.rcsizer  as rcs

class InputPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.filepath1 = ""
        self.filepath2 = ""

        sizer = rcs.RowColSizer()

        text = u"请先选择第一个文件，然后选择第二个文件，最后点击生成按钮"

        sizer.Add(wx.StaticText(self, -1, text), row=1, col=2, colspan=4)

        textctl = wx.TextCtrl(self, style=wx.TE_READONLY)
        self.textctl1 = textctl
        sizer.Add(textctl, row=3, col=2, colspan=3, flag=wx.EXPAND)
        button = wx.Button(self, wx.ID_ANY, u"选择第一个文件")
        self.Bind(wx.EVT_BUTTON, self.OnButton1, button)
        sizer.Add(button, row=3, col=5, colspan=1)

        textctl = wx.TextCtrl(self, style=wx.TE_READONLY)
        self.textctl2 = textctl
        sizer.Add(textctl, row=5, col=2, colspan=3, flag=wx.EXPAND)
        button = wx.Button(self, wx.ID_ANY, u"选择第二个文件")
        self.Bind(wx.EVT_BUTTON, self.OnButton2, button)
        sizer.Add(button, row=5, col=5, colspan=1)

        button = wx.Button(self, wx.ID_ANY, u"生成")
        self.Bind(wx.EVT_BUTTON, self.OnButtonGenerate, button)
        sizer.Add(button, row=7, col=2, colspan=4, flag=wx.EXPAND)

        self.SetSizer(sizer)

    def OnButton1(self, event):
        dlg = wx.FileDialog(
            self, message="打开文件", defaultDir=os.getcwd(),
            defaultFile="", wildcard="All files (*.*)|*.*", style=wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filepath1 = dlg.GetPath()
            self.textctl1.WriteText(self.filepath1)
        dlg.Destroy()

    def OnButton2(self, event):
        dlg = wx.FileDialog(
            self, message="打开文件", defaultDir=os.getcwd(),
            defaultFile="", wildcard="All files (*.*)|*.*", style=wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filepath2 = dlg.GetPath()
            self.textctl2.WriteText(self.filepath2)
        dlg.Destroy()

    def OnButtonGenerate(self, event):
        if self.filepath1 == "" or self.filepath2 == "":
            dlg = wx.MessageDialog(self,
                                   u'你少选了至少一个文件！！！！',
                                   u'注意啦', wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
            return

        resultFileName = opExcel.main(self.filepath1,self.filepath2)
        dlg = wx.MessageDialog(self,
                               u'生成完毕！\n输出文件：' + resultFileName,
                               u'注意啦', wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()

class GUI(object):
    def run(self):
        app = wx.App(False)
        frame = wx.Frame(None, wx.ID_ANY, "opExcel", size=(420, 300))
        frame.CreateStatusBar()

        panel = InputPanel(frame)
        frame.Show(True)
        app.MainLoop()

if __name__ == '__main__':
    gui = GUI()
    gui.run()



