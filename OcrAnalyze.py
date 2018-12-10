import FrameWork
import wx
import os
import datetime
import time
import win32api
from pdf2jpg import pdf2jpg
import xlrd
import xlwt
from xlutils.copy import copy
from pytesseract import *
import re
import threading
import analyzeStr
from multiprocessing import Process
import shutil

class showFrame(FrameWork.MyFrame1,analyzeStr.AnalyzeStr):
    #Start with establish the frame and initialize.
    def __init__(self,parent):
        FrameWork.MyFrame1.__init__(self, parent)
        self.choose_input.Bind(wx.EVT_BUTTON,self.chooseIn)
        self.BothStart.Bind(wx.EVT_BUTTON,self.start)
        self.choose_input.Bind(wx.EVT_ENTER_WINDOW, self.showChoose)
        self.BothStart.Bind(wx.EVT_ENTER_WINDOW, self.showStart)
        self.process=0
        self.totalmission=0
    #__init__ 初始化程序，绑定框架，绑定按键与事件
    #Initialize the program, bind the frame and buttons.

    def showChoose(self,event):
        print('准备工作提示\n'+'你需要将要录入的S号文件复制并装入一个文件夹中\n然后再从"选择文件夹"的按钮中选择这个文件夹：\n---------------------------------------------------------------')
    #showChoose当鼠标移至按键上面的时候，在命令行显示提示
    #When the mouse is on this btn, show attention at the command line.
    def showStart(self,event):
        print('开始工作提示\n' + '当您完成准备工作后即可点击此键开始运行\n识别过程中的进展你可以在这个窗口进行监督：\n---------------------------------------------------------------')
    # showStart当鼠标移至按键上面的时候，在命令行显示提示
    #When the mouse is on this btn, show attention at the command line.
    def chooseIn(self,event):
        #chooseIn选择需要扫描的文件夹，并选中里面筛选后的文件
        #Select the folder you want to deal with.This will help with to filter the pdf and pics.
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        self.inputpath=''
        if dlg.ShowModal() == wx.ID_OK:
            self.inputpath=dlg.GetPath()
            print('您选择的文件夹： '+self.inputpath)  # 文件夹路径
        dlg.Destroy()
        count=0
        for root,dirs,files in os.walk(self.inputpath):
            dirs.sort()
            self.total_dir=dirs
            count+=1
            break
        print('dir')
        print(self.total_dir)
        self.eachFile()
        self.eachTIF()
        wx.MessageBox('您已可以点击开始键运行程序~')
        self.m_gauge2.SetRange(self.totalmission)

    def eachTIF(self):
        #Filter the tif files. Only the one respect the requirments are available.
        #筛选TIF文件，只有满足需求的才能被使用。
        pathDir = os.listdir(self.inputpath)
        self.each_TIF=[]
        self.eachPerTIF=[]
        self.EngineerName={}
        self.newdir = os.listdir(self.inputpath)
        self.eachperInv=[]
        self.eachInv=[]
        for root, dirs, files in os.walk(self.inputpath):
            for i in files:
                child=('%s\%s'%(root,i))
                if (os.path.splitext(child)[1] == ".TIF" or os.path.splitext(child)[1] == ".tif")and self.anaTitle(i):
                    self.eachPerTIF.append(i)
                    self.each_TIF.append(child)
                if os.path.splitext(child)[1] == ".DOC" or os.path.splitext(child)[1] == ".doc":
                    if i[:-4]!=r'调单协议' and i[:-4]!=r'Pre_Contact_Statement' and re.findall(u'报关委托书',i)==[] and not(self.anaTitle(i)):
                        self.EngineerName[root[-7:]]=i[:-4]
                if (os.path.splitext(child)[1] == ".TIF" or os.path.splitext(child)[1] == ".tif")and (not(self.anaTitle(i))):
                    self.eachperInv.append(i)
                    self.eachInv.append(child)
        self.totalmission += (len(self.eachPerTIF))
        self.newdir = os.listdir(self.inputpath)
        strInputPath = ''
        for i in self.each_TIF:
            strInputPath += str(i) + '\n'

    def start(self,event):
        #This part is the main loop.
        #这一部分是主要程序的循环。
        wildcard = u"Access 数据库 (*.accdb)|*.accdb"
        dlg = wx.FileDialog(self, message=u"选择文件",
                            defaultDir=os.getcwd(),
                            defaultFile="",
                            wildcard=wildcard,
                            style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.DBfile = (dlg.GetPaths())[0]
            print(self.DBfile)
        dlg.Destroy()
        #First is to choose the folder to OCR.
        starttime = datetime.datetime.now()
        self.inv_list=[]
        self.pdf_ocr()
        self.tif_ocr()
        #Second is to OCR these files.
        tifList=self.checkSupply(self.tiflist)
        pdfList=self.checkSupply(self.pdflist)
        end_list=pdfList+tifList
        #Then is to analyze and get the info which meets the requirments.
        ccount=0
        for i in end_list:
            appending='none'
            for kkkk in self.inv_list:
                if kkkk[0]==i[0] and i[6]=='none':
                    appending=kkkk[1]
            end_list[ccount][6]=appending
            ccount+=1
        keys_ENGI=list(self.EngineerName.keys())
        a=0
        for i in end_list:
            i_a=i[0].lower()
            i[0]=i_a
            for n in keys_ENGI:
                if i[0]==n:
                    end_list[a][2]=self.EngineerName[n]
            a+=1
        repeated_=[]
        for i in range(len(end_list)-1):
            if end_list[i][0]==end_list[i+1][0]:
                for b in range((len(end_list[i]))):
                    if end_list[i][b] =='none':
                        end_list[i][b]=end_list[i+1][b]
                if end_list[i][2]=='ACEWAY':
                    repeated_.append(i)
                else:
                    repeated_.append(i+1)
        c_n=0
        countcc=1
        for z in repeated_:
            if c_n>0:
                z=z-countcc
                countcc+=1
            end_list.pop(z)
            c_n+=1
        the_end_list=[]
        for each in end_list:
            tuple_one = tuple(each)
            the_end_list.append(tuple_one)
        def takeFirst(elem):
            return elem[0]
        the_end_list.sort(key=takeFirst)
        print(the_end_list)
        endtime = datetime.datetime.now()
        print(str((endtime - starttime).seconds) + ' seconds')
        nowtime_ = time.localtime()
        nowtitle = str(nowtime_.tm_mon) + "." + str(nowtime_.tm_mday) + "_" + str(nowtime_.tm_hour) + "." + str(
            nowtime_.tm_min)
        excelname = 'import list '+nowtitle+'.xls'
        self.write2excel(the_end_list,excelname)
        self.accdb_(the_end_list,self.DBfile)

    def pdf_ocr(self):
        self.pdfstring = ''
        self.pdflist=[]
        termlist = ['EXW', 'CIP', 'CIF', 'FCA', 'FOB', 'DDU', 'DAP']
        inv_list = []
        outputpath = self.inputpath
        for k in self.eachpdfinv:
            inputpath = str(k)
            for n in self.newdir:
                s_part = re.findall('\d+', n)[0]
                s_tr = 's' + s_part
                if s_part in k:
                    s_num = s_tr
                    break
            invres=(pdf2jpg.convert_pdf2jpg(inputpath, outputpath, pages="ALL"))
            invadd=invres[0]['output_jpgfiles']
            term = 'none'
            limit = 0
            for kkk in invadd:
                text = pytesseract.image_to_string(kkk)
                for a_min in termlist:
                    if a_min in text:
                        term = a_min
                        break
                if limit>=2:
                    break
                limit+=1
            ll=[]
            ll.append(s_num)
            ll.append(term)
            self.inv_list.append(ll)
        a = 0
        b = []
        count_ = 0
        item = 0
        if self.eachPDF!=[]:
            outputpath = self.inputpath
            result = []
            for i in self.eachPDF:
                inputpath = str(i)
                pack='none'
                weight='none'
                for n in self.newdir:
                    num_part=re.findall(r'\d+', n)[0]
                    s_tr = 's' + num_part
                    if num_part in i:
                        s_number = s_tr
                        break
                invres =(pdf2jpg.convert_pdf2jpg(inputpath, outputpath, pages="ALL"))

                invadd = invres[0]['output_jpgfiles']
                count_ += 1
                hwb_real = self.analyzeNum(self.eachPerName[a])
                forwardername=self.checkforwarder(self.eachPerName[a])
                textChecked=[]
                limit = 0
                for n in invadd:
                    k = n
                    text = pytesseract.image_to_string(k)
                    for i in self.checkTXT(text):
                        textChecked.append(i)
                    if re.findall(r"\d+\s\d+k", text, re.I) != []:
                        p_w = re.findall(r"\d+\s\d+k", text, re.I)[0]
                        pack = str(re.findall(r"\d+", p_w, re.I)[0])
                        weight = str(re.findall(r"\d+", p_w, re.I)[1])
                    elif re.findall(r"\d+\s\d+.{3,4}K", text, re.I) != []:
                        p_w = re.findall(r"\d+\s\d+.{3,4}K", text, re.I)[0]
                        pack = str(re.findall(r"\d+", p_w, re.I)[0])
                        weight = ''
                        if len(re.findall(r"\d+", p_w, re.I)) == 2:
                            weight += str(re.findall(r"\d+", p_w, re.I)[1])
                        else:
                            weight += str(re.findall(r"\d+", p_w, re.I)[1])
                            weight += '.'
                            weight += str(re.findall(r"\d+", p_w, re.I)[2])
                    if limit >= 2:
                        break
                    limit += 1
                textChecked.append(forwardername)
                b.append(s_number + ": " + str(hwb_real) + str(textChecked))
                self.pdflist.append([s_number, hwb_real, textChecked,pack,weight])
                item+=1
                a+=1
                print('PDF已完成 '+str(item)+' 单'+'\n运单号为： '+self.eachPerName[a-1])
            self.pdfstring = str(b)
        else:
            pass

    def tif_ocr(self):
        a=0
        b=[]
        s_num='none-s'
        self.tiflist=[]
        item=0
        termlist = ['EXW', 'CIP', 'CIF', 'FCA', 'FOB', 'DDU', 'DAP']
        inv_list = []
        for k in self.eachInv:
            for n in self.newdir:
                if n in k:
                    s_num = n
                    break
            text = pytesseract.image_to_string(str(k))
            term='none'
            for a_min in termlist:
                if a_min in text:
                    term = a_min
                    break
            ll = []
            ll.append(s_num)
            ll.append(term)
            self.inv_list.append(ll)
        for i in self.each_TIF:
            if ('Inv'or'inv'or'INV' or 'Inv1') in str(i):
                pass
            else:
                self.process += 1
                self.m_gauge2.SetValue(self.process)

                for n in self.newdir:
                    if n in i:
                        s_num = n
                        break
                text = pytesseract.image_to_string(str(i))
                hwb_real = self.analyzeNum(self.eachPerTIF[a])
                textChecked = self.checkTXT(text)
                textChecked.append(self.checkforwarder(self.eachPerTIF[a]))
                b.append(s_num + str(hwb_real) + str(textChecked))
                if re.findall(r"\d+\s\d+k", text, re.I) != []:
                    p_w =re.findall(r"\d+\s\d+k", text, re.I)[0]
                    pack = str(re.findall(r"\d+", p_w, re.I)[0])
                    weight = str(re.findall(r"\d+", p_w, re.I)[1])
                elif  re.findall(r"\d+\s\d+.{3,4}K", text, re.I) != []:
                    p_w =  re.findall(r"\d+\s\d+.{3,4}K", text, re.I)[0]
                    pack = str(re.findall(r"\d+", p_w, re.I)[0])
                    weight=''
                    if len(re.findall(r"\d+", p_w, re.I))==2:
                        weight+=str(re.findall(r"\d+", p_w, re.I)[1])
                    else:
                        weight+=str(re.findall(r"\d+", p_w, re.I)[1])
                        weight+='.'
                        weight+=str(re.findall(r"\d+", p_w, re.I)[2])

                else:
                    pack='none'
                    weight = 'none'
                self.tiflist.append([s_num, hwb_real, textChecked,pack,weight])
                item+=1
                print('TIF图片已完成 ' + str(item) + ' 单'+'\n运单号为： '+self.eachPerTIF[a-1])
        self.tifstring=str(b)

    def eachFile(self):
        pathDir = os.listdir(self.inputpath)
        self.eachPDF=[]
        self.eachPerName=[]
        self.eachpdfinv=[]
        self.eachperpdfinv=[]
        rename=0
        for root, dirs, files in os.walk(self.inputpath):
            for i in files:
                child=('%s\%s'%(root,i))
                if (os.path.splitext(child)[1] == ".PDF" or os.path.splitext(child)[1] == ".pdf" )and self.anaTitle(i):
                    self.eachPerName.append(i)
                    self.eachPDF.append(child)
                if (os.path.splitext(child)[1] == ".PDF" or os.path.splitext(child)[1] == ".pdf" )and (not (self.anaTitle(i))) and self.contract(i):
                    newi=r'\inv'+str(rename)+'.pdf'
                    newname = root + newi
                    shutil.copyfile(child, newname)
                    time.sleep(0.5)
                    self.eachperpdfinv.append(newi)
                    self.eachpdfinv.append(newname)
                    rename+=1
        self.totalmission +=(len(self.eachPerName))
        strInputPath=''
        for i in self.eachPDF:
            strInputPath += str(i)+'\n'

    def pdfTojpg(self,event):
        outputpath=self.outputpath
        result=[]
        for i in self.eachPDF:
            inputpath=str(i)
            result.append(pdf2jpg.convert_pdf2jpg(inputpath, outputpath, pages="0"))
        a=0
        b=[]
        for n in result:
            k=str(str(n[0]['output_jpgfiles'][0]))
            text = pytesseract.image_to_string(k)
            hwb_demo = self.getNum(self.eachPerTIF[a])
            a += 1
            b.append(str(self.analyze(hwb_demo,text))+str(self.checkTXT(text)))
            
try:
    app = wx.App(False)
    frame =  showFrame(None)
    frame.Show(True)
    app.MainLoop()
except:
    pass
