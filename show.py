import wx
import win32com.client
from PIL import Image
import os
import shutil
import re


class Mywinttp(wx.Frame):
    def __init__(self,parent,title):
        super(Mywinttp,self).__init__(parent,title=title,size=(500,200))
        self.icon = wx.Icon('st.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(self.icon)

        self.pptfiles=None
        self.InitUI()
        self.Centre()
        self.Show()

    def InitUI(self):
        panel = wx.Panel(self)
        sizer = wx.GridBagSizer(0, 0)

        self.text = wx.StaticText(panel, label="选择PPT文件:")
        sizer.Add(self.text, pos=(0, 1), flag=wx.ALL | wx.ALIGN_CENTER, border=5)

        self.tc = wx.TextCtrl(panel,size = (250,20))
        sizer.Add(self.tc, pos=(0, 2), span=(0, 2), flag=wx.EXPAND |wx.ALIGN_CENTER | wx.ALL , border=5)

        self.tbt = wx.Button(panel, label='浏览')
        sizer.Add(self.tbt,pos=(0,4),span=(0,3),flag=wx.EXPAND | wx.ALL, border=5)
        self.Bind(wx.EVT_BUTTON, self.OnclickFile,self.tbt)

        self.text1 = wx.StaticText(panel, label="输出目录:")
        sizer.Add(self.text1, pos=(1, 1), flag=wx.ALL | wx.ALIGN_CENTER, border=5)

        self.tc1 = wx.TextCtrl(panel, size=(250, 20))
        sizer.Add(self.tc1, pos=(1, 2), span=(1, 2), flag=wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, border=5)

        self.tbt1 = wx.Button(panel, label='浏览')
        sizer.Add(self.tbt1, pos=(1, 4), span=(1, 3), flag=wx.EXPAND | wx.ALL, border=5)
        self.Bind(wx.EVT_BUTTON, self.OnclickDir, self.tbt1)

        self.tbt2 = wx.Button(panel, label='转化',size=(50,30))
        sizer.Add(self.tbt2, pos=(3, 2), span=(3, 2), flag=wx.EXPAND | wx.ALL, border=5)
        self.Bind(wx.EVT_BUTTON, self.clickZh, self.tbt2)

        panel.SetSizerAndFit(sizer)

    def OnclickFile(self, e):
        dlg = wx.DirDialog(self, "选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            url = dlg.GetPath()
            self.tc.write(url)
        dlg.Destroy()
        # wildcard = "PPT Files (*.pptx)|*.pptx|PPT Files (*.ppt)|*.ppt|所有文件 (*.*)|*.*"
        # dlg = wx.FileDialog(self, "选择要转化的PPT文件", os.getcwd(), "", wildcard, wx.FD_OPEN | wx.FD_MULTIPLE)
        # if dlg.ShowModal() == wx.ID_OK:
        #     url = dlg.GetPaths()
        #     self.tc.SetValue('')
        #     self.pptfiles = url
        #     self.tc.write((';').join(url))
        # dlg.Destroy()
    def OnclickDir(self,e):
        dlg = wx.DirDialog(self, "选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            url = dlg.GetPath()
            self.tc1.write(url)
        dlg.Destroy()

    def pptSave(self,pptfile,outdir,fileName):
        outDir = outdir+'/'+fileName
        power = win32com.client.Dispatch('PowerPoint.Application')
        power.Visible=1
        ppt = power.Presentations.Open(pptfile)
        try:
            ppt.SaveAs(outDir + '.jpg', 17)
            power.Quit()
        except BaseException:
            wx.MessageBox(fileName+"文件错误", "Message", wx.OK | wx.ICON_INFORMATION)


    def clickZh(self,e):
        pptfiles = self.tc.GetValue()
        outdir = self.tc1.GetValue()
        if not pptfiles.strip():
            wx.MessageBox("请选择ppt文件", "Message", wx.OK | wx.ICON_INFORMATION)
            return False
        if not outdir.strip():
            wx.MessageBox("请选择输出目录", "Message", wx.OK | wx.ICON_INFORMATION)
            return False
        self.tbt2.Enable(False)
        self.tbt2.SetLabel("正在转化中...")
        for i in (fns for fns in os.listdir(pptfiles.strip()) if fns.endswith(('.ppt', '.pptx'))):
            i=pptfiles+'\\'+i
            fileName=os.path.split(i)
            fileName=os.path.splitext(fileName[1])
            fileName=fileName[0].strip()
            self.pptSave(i, outdir,fileName)
            self.ppt2jpg(outdir,fileName)

        self.tbt2.Enable(True)
        self.tbt2.SetLabel("转化")
        wx.MessageBox("转化成功", "Message", wx.OK | wx.ICON_INFORMATION)


    #ppt转化成jpg
    def ppt2jpg(self,outdir,fileName):
        configs = self.readConfig();
        im = Image.open(configs['tem'])
        outDirTem = outdir + '/' + fileName
        n=1
        w1=int(configs['w1'])
        h1=int(configs['h1'])
        x1=int(configs['x1'])
        y1=int(configs['y1'])
        w2=int(configs['w2'])
        h2=int(configs['h2'])
        x2=int(configs['x2'])
        y2=int(configs['y2'])
        line=int(configs['line'])
        max=int(configs['max'])
        piclist = os.listdir(outDirTem)
        piclist.sort(key = lambda i:int(re.search(r'(\d+)',i).group(0)))
        for fn in (fns for fns in piclist if fns.endswith(('.jpg', '.JPG'))):
            if n>max:
                break
            if n==1:
                box = (x1, y1, x1+w1, y1+h1)
                resize_x=(w1, h1)
            elif (n % 2) == 0:
                if n==2:
                    linec=0
                else:
                    linec=line

                cy = y2 + ((h2+ linec) * (int(n/2) - 1))
                cy2 = y2 + ((h2+ linec) * (int(n/2) - 1)) + h2

                box = (x2,cy  , x2 + w2 ,cy2 )
                resize_x = (w2, h2)
            else:
                if n==3:
                    linec=0
                else:
                    linec=line
                cy = y2 + ((h2+ linec) * (int((n-1)/2)-1))
                cy2 = y2 + ((h2+ linec) * (int((n-1)/2)-1))+h2
                box = (x2+w2+line, cy, x2+w2+w2+line, cy2)
                resize_x = (w2, h2)
            fx = outDirTem + '\\' + fn
            imx = Image.open(fx)
            imx = imx.resize(resize_x,Image.ANTIALIAS)
            im.paste(imx, box)
            n=n+1
        width = int(im.size[0])
        height=int((n-1)/2)*(h2+line)+y2
        box_end  = (0,0,width,height)
        im = im.crop(box_end)
        im.save(outdir+"\\"+fileName+".jpg",quality = 90)
        shutil.rmtree(outDirTem)


    def readConfig(self):
        file = open('c.txt', 'r')
        cstr = file.read()
        config_str = {}
        for item in cstr.split("\n"):
            cstr = item.split(":")
            config_str[cstr[0]] = cstr[1]
        file.close()
        return config_str

ex = wx.App()
Mywinttp(None,"ppt转jpg样机-图素网")
ex.MainLoop()

