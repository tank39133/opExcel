#coding=utf8

__author__ = 'FuQiang'

import os
import os.path
import opExcel
import pprint

import wx
import  wx.lib.rcsizer  as rcs

class InputPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.filepaths = []

        sizer = rcs.RowColSizer()

        text = u"请添加文件到下方的列表，然后点击生成。"

        sizer.Add(wx.StaticText(self, -1, text), row=1, col=2, colspan=10)

        self.listBox = wx.ListBox(self, size=(550, 300), style=wx.LB_HSCROLL)
        sizer.Add(self.listBox, row=3, col=2, colspan=9, rowspan=8, flag=wx.EXPAND)

        button = wx.Button(self, wx.ID_ANY, u"添加")
        self.Bind(wx.EVT_BUTTON, self.OnAdd, button)
        sizer.Add(button, row=3, col=12)

        button = wx.Button(self, wx.ID_ANY, u"清除")
        self.Bind(wx.EVT_BUTTON, self.OnClear, button)
        sizer.Add(button, row=4, col=12)

        button = wx.Button(self, wx.ID_ANY, u"生成")
        self.Bind(wx.EVT_BUTTON, self.OnButtonGenerate, button)
        sizer.Add(button, row=8, col=12)

        self.SetSizer(sizer)

    def OnAdd(self, event):
        dlg = wx.FileDialog(
            self, message="打开多个文件", defaultDir=os.getcwd(),
            defaultFile="", wildcard="All files (*.*)|*.*", style=wx.OPEN | wx.MULTIPLE)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            for path in paths:
                if path not in self.filepaths:
                    self.filepaths.append(path)

            self.listBox.Clear()
            for path in self.filepaths:
                self.listBox.Append(path)

        dlg.Destroy()

    def OnClear(self, event):
        self.listBox.Clear()
        self.filepaths = []

    def OnButtonGenerate(self, event):
        if len(self.filepaths) == 0:
            dlg = wx.MessageDialog(self,
                                   u'你应该至少选择一个文件！！！！',
                                   u'注意啦', wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
            return

        pprint.pprint(self.filepaths)


        dlg = wx.MessageDialog(self,
                               u'生成完毕！\n输出文件：',
                               u'注意啦', wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()

class GUI(object):
    def run(self):
        app = wx.App(False)
        frame = wx.Frame(None, wx.ID_ANY, "opExcel", size=(720, 460))
        frame.CreateStatusBar()

        panel = InputPanel(frame)
        frame.Show(True)
        app.MainLoop()

if __name__ == '__main__':
    gui = GUI()
    gui.run()



