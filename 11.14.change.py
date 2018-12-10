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
    def __init__(self,parent):
        FrameWork.MyFrame1.__init__(self, parent)
        self.choose_input.Bind(wx.EVT_BUTTON,self.chooseIn)
        # self.choose_output.Bind(wx.EVT_BUTTON,self.chooseOut)
        # self.Start.Bind(wx.EVT_BUTTON,self.pdfTojpg)
        # self.tif2ocr.Bind(wx.EVT_BUTTON,self.T2O)
        self.BothStart.Bind(wx.EVT_BUTTON,self.start)
        self.choose_input.Bind(wx.EVT_ENTER_WINDOW, self.showChoose)
        self.BothStart.Bind(wx.EVT_ENTER_WINDOW, self.showStart)
        # self.ChooseList.Bind(wx.EVT_BUTTON,self.chooseList)
        self.process=0
        self.totalmission=0
    #__init__初始化程序，绑定框架，绑定按键与事件

    def showChoose(self,event):
        print('准备工作提示\n'+'你需要将要录入的S号文件复制并装入一个文件夹中\n然后再从"选择文件夹"的按钮中选择这个文件夹：\n---------------------------------------------------------------')
    #showChoose当鼠标移至按键上面的时候，在命令行显示提示
    def showStart(self,event):
        print('开始工作提示\n' + '当您完成准备工作后即可点击此键开始运行\n识别过程中的进展你可以在这个窗口进行监督：\n---------------------------------------------------------------')
    # showStart当鼠标移至按键上面的时候，在命令行显示提示
    def chooseIn(self,event):
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        self.inputpath=''
        if dlg.ShowModal() == wx.ID_OK:
            self.inputpath=dlg.GetPath()

            print('您选择的文件夹： '+self.inputpath)  # 文件夹路径
        dlg.Destroy()
        count=0
        for root,dirs,files in os.walk(self.inputpath):
            # print(root)

            # print(dirs)
            dirs.sort()
            self.total_dir=dirs
            count+=1
            break
        print('dir')
        print(self.total_dir)
        # self.start_dir=total_dir[0]
        # self.end_dir=total_dir[len(total_dir)-1]

        # if total_dir[0]>total_dir[len(total_dir)-1]:
        #     self.start_dir=(total_dir[len(total_dir)-1]).split('s')[1]
        #     self.end_dir=total_dir[0].split('s')[1]
        # else:
        #     self.start_dir = total_dir[0].split('s')[1]
        #     self.end_dir = total_dir[len(total_dir) - 1].split('s')[1]
        # print(self.start_dir)
        # print(self.end_dir)

            # print(files)
        self.eachFile()
        self.eachTIF()
        wx.MessageBox('您已可以点击开始键运行程序~')
        # print(os.getcwd())

        self.m_gauge2.SetRange(self.totalmission)
    #chooseIn选择需要扫描的文件夹，并选中里面筛选后的文件

    def chooseList(self,event):
        wildcard="Text Files (*.xls)|*.xls|""EXCEL XLSX (*.xlsx)|*.xlsx"
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd(), "", wildcard)
        if dlg.ShowModal() == wx.ID_OK:
            workbook = xlrd.open_workbook(dlg.GetPath(),'formatting_info=True')
        dlg.Destroy()
        sheet=workbook.sheet_by_name('Tabelle1')
        rowNum = sheet.nrows
        colNum = sheet.ncols
        data=[]

        # newbook = copy(workbook)
        #
        # newsheet=newbook.get_sheet(0)
        # str = 'hehe'
        # newsheet.write(rowNum, 0, str)
        # # 覆盖保存
        # newbook.save('gg.xls')
    #

    def chooseOut(self,event):
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        self.outputpath=''
        if dlg.ShowModal() == wx.ID_OK:
            self.outputpath=dlg.GetPath()
            print(self.outputpath)  # 文件夹路径
        dlg.Destroy()

    def eachTIF(self):
        pathDir = os.listdir(self.inputpath)
        self.each_TIF=[]
        self.eachPerTIF=[]
        self.EngineerName={}
        self.newdir = os.listdir(self.inputpath)
        self.eachperInv=[]
        self.eachInv=[]
        for root, dirs, files in os.walk(self.inputpath):
            # print(root)
            for i in files:
                child=('%s\%s'%(root,i))
                if (os.path.splitext(child)[1] == ".TIF" or os.path.splitext(child)[1] == ".tif")and self.anaTitle(i):
                    print(self.analyzeNum(i))
                    self.eachPerTIF.append(i)
                    self.each_TIF.append(child)
                if os.path.splitext(child)[1] == ".DOC" or os.path.splitext(child)[1] == ".doc":
                    if i[:-4]!=r'调单协议' and i[:-4]!=r'Pre_Contact_Statement' and re.findall(u'报关委托书',i)==[] and not(self.anaTitle(i)):
                        self.EngineerName[root[-7:]]=i[:-4]
                if (os.path.splitext(child)[1] == ".TIF" or os.path.splitext(child)[1] == ".tif")and (not(self.anaTitle(i))):
                    self.eachperInv.append(i)
                    self.eachInv.append(child)
        # print('eachperInv:')
        # print(self.eachperInv)
        # print('eachInv')
        # print(self.eachInv)


        self.totalmission += (len(self.eachPerTIF))
        # print(self.EngineerName)


        self.newdir = os.listdir(self.inputpath)
        # print(self.newdir)
        strInputPath = ''
        for i in self.each_TIF:
            strInputPath += str(i) + '\n'

    def importFromxls(self,event):
        pass


    def T2O(self,event):
        if self.each_TIF!=[]:
            a = 0
            b = []
            for i in self.each_TIF:
                text = pytesseract.image_to_string(str(i))
                hwb_demo = self.getNum(self.eachPerTIF[a])
                print(hwb_demo)
                b.append(str(self.analyze(hwb_demo, text)) + str(self.checkTXT(text)))
                print(i)
                f = open(i + ".txt", 'w', encoding='utf-8')
                f.write(text)
                f.close()
                a += 1
            f = open(self.inputpath + r"\result.txt", 'w', encoding='utf-8')
            f.write(str(b))
            f.close()
            print('Over-------------------------------------')
            pass
        else:
            pass

    def start(self,event):
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
        starttime = datetime.datetime.now()
        # pstring=''
        self.inv_list=[]
        self.pdf_ocr()
        self.tif_ocr()
        print(self.inv_list)

        tifList=self.checkSupply(self.tiflist)
        pdfList=self.checkSupply(self.pdflist)
        end_list=pdfList+tifList
        print(end_list)
        ccount=0
        for i in end_list:
            appending='none'
            for kkkk in self.inv_list:
                if kkkk[0]==i[0] and i[6]=='none':
                    appending=kkkk[1]
            # end_list[ccount].append(appending)
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
        # print('before:')
        # print(end_list)
        for i in range(len(end_list)-1):
            if end_list[i][0]==end_list[i+1][0]:
                print('there is one')
                for b in range((len(end_list[i]))):
                    if end_list[i][b] =='none':
                        end_list[i][b]=end_list[i+1][b]
                if end_list[i][2]=='ACEWAY':
                    repeated_.append(i)
                else:
                    repeated_.append(i+1)
        print(end_list)
        print(repeated_)
        c_n=0
        countcc=1
        for z in repeated_:
            # print('i will del this')
            # print(z)
            if c_n>0:
                z=z-countcc
                countcc+=1
            # del end_list[z]

            end_list.pop(z)
            c_n+=1

        # print('later:')
        # print(end_list)
        the_end_list=[]
        print(end_list)
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
        # self.adb_(the_end_list)
        self.accdb_(the_end_list,self.DBfile)
        # msgDLG=wx.MessageBox('所有的数据已经识别完成并且导入到了此EXCEL表格里：'+excelname,'成功！')
        #-------------------------------------------------------------------------------------------------------------------------------------------------
        # result = wx.MessageBox('所有的数据已经识别完成并且导入到了此EXCEL表格里：'+excelname+'\n当您关闭这个窗口或者点击OK时， 会自动打开此EXCEL文件~','成功！',
        #                        wx.OK | wx.CANCEL | wx.ICON_EXCLAMATION)
        # if result == wx.OK:
        #     win32api.ShellExecute(0, 'open',excelname+'.xls' , '', '', 1)
        #     wx.Exit()
        # elif result == wx.CANCEL:
        #     wx.Exit()

        #=========================================================================================================================================================
        # wx.Exit()

    def pdf_ocr(self):
        self.pdfstring = ''
        self.pdflist=[]
        termlist = ['EXW', 'CIP', 'CIF', 'FCA', 'FOB', 'DDU', 'DAP']
        inv_list = []
        outputpath = self.inputpath
        # invres=[]

        for k in self.eachpdfinv:
            inputpath = str(k)
            # s_tr='s'+re.findall('\d+',k)[0]
            # print(s_tr)
            for n in self.newdir:
                s_part = re.findall('\d+', n)[0]
                s_tr = 's' + s_part
                print(s_tr)
                if s_part in k:

                    s_num = s_tr
                    break
            invres=(pdf2jpg.convert_pdf2jpg(inputpath, outputpath, pages="ALL"))
            # text = pytesseract.image_to_string(str(k))
            invadd=invres[0]['output_jpgfiles']
            print(invres)
            term = 'none'
            limit = 0
            for kkk in invadd:
                #add=str(str(invres[0]['output_jpgfiles'][0]))
                text = pytesseract.image_to_string(kkk)
                for a_min in termlist:
                    if a_min in text:
                        term = a_min
                        break
                if limit>=2:
                    break
                limit+=1
            # f = open(add + ".txt", 'w', encoding='utf-8')
            # f.write(text)
            # f.close()

            ll=[]
            ll.append(s_num)
            ll.append(term)
            self.inv_list.append(ll)

        # print(inv_list)
        a = 0
        b = []
        count_ = 0
        item = 0
        if self.eachPDF!=[]:
            outputpath = self.inputpath
            result = []
            for i in self.eachPDF:
                inputpath = str(i)
                # print(i)
                # s_tr = 's' + re.findall(r'\d+', i)[0]
                # print(s_tr)
                pack='none'
                weight='none'
                for n in self.newdir:
                    num_part=re.findall(r'\d+', n)[0]
                    s_tr = 's' + num_part
                    print(s_tr)
                    if num_part in i:
                        s_number = s_tr
                        break
                invres =(pdf2jpg.convert_pdf2jpg(inputpath, outputpath, pages="ALL"))

                invadd = invres[0]['output_jpgfiles']
                # print(invres)
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
                        # print(re.findall(r"\d+\s\d+k", text, re.I)[0])
                        p_w = re.findall(r"\d+\s\d+k", text, re.I)[0]
                        pack = str(re.findall(r"\d+", p_w, re.I)[0])
                        weight = str(re.findall(r"\d+", p_w, re.I)[1])
                    elif re.findall(r"\d+\s\d+.{3,4}K", text, re.I) != []:
                        p_w = re.findall(r"\d+\s\d+.{3,4}K", text, re.I)[0]
                        pack = str(re.findall(r"\d+", p_w, re.I)[0])
                        # weight = str(re.findall(r"\d+", p_w, re.I)[1])
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
                print(textChecked)
                b.append(s_number + ": " + str(hwb_real) + str(textChecked))
            #         if re.findall(r"\d+\s\d+k", text, re.I) != []:
            # # print(re.findall(r"\d+\s\d+k", text, re.I)[0])
            #             p_w = re.findall(r"\d+\s\d+k", text, re.I)[0]
            #             pack = str(re.findall(r"\d+", p_w, re.I)[0])
            #             weight = str(re.findall(r"\d+", p_w, re.I)[1])
            #         elif re.findall(r"\d+\s\d+.{3,4}K", text, re.I) != []:
            #             p_w = re.findall(r"\d+\s\d+.{3,4}K", text, re.I)[0]
            #             pack = str(re.findall(r"\d+", p_w, re.I)[0])
            # # weight = str(re.findall(r"\d+", p_w, re.I)[1])
            #             weight = ''
            #             if len(re.findall(r"\d+", p_w, re.I)) == 2:
            #                 weight += str(re.findall(r"\d+", p_w, re.I)[1])
            #             else:
            #                 weight += str(re.findall(r"\d+", p_w, re.I)[1])
            #                 weight += '.'
            #                 weight += str(re.findall(r"\d+", p_w, re.I)[2])
            #
            #         else:
            # # print('here!')
            #             pack = 'none'
            #             weight = 'none'


                # if re.findall(r"\d+\s\d+.{3,4}KG", text, re.I) != []:
                #     p_w = re.findall(r"\d+\s\d+.{3,4}KG", text, re.I)[0]
                #     pack = str(re.findall(r"\d+", p_w, re.I)[0])
                #     weight = str(re.findall(r"\d+", p_w, re.I)[1])


                self.pdflist.append([s_number, hwb_real, textChecked,pack,weight])
                item+=1
                a+=1
                print('PDF已完成 '+str(item)+' 单'+'\n运单号为： '+self.eachPerName[a-1])
            self.pdfstring = str(b)
            # print(self.pdfstring)

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
            f = open(s_num + ".txt", 'w', encoding='utf-8')
            f.write(text)
            f.close()
            term='none'
            for a_min in termlist:
                if a_min in text:
                    term = a_min
                    break
            ll = []
            ll.append(s_num)
            ll.append(term)
            self.inv_list.append(ll)
        # print(inv_list)
        for i in self.each_TIF:
            # print(str(i))
            if ('Inv'or'inv'or'INV' or 'Inv1') in str(i):
                pass
                # print('Not This')
            else:
                self.process += 1
                self.m_gauge2.SetValue(self.process)

                for n in self.newdir:
                    if n in i:
                        s_num = n
                        break
                # wx.MessageBox('three')
                text = pytesseract.image_to_string(str(i))
                # wx.MessageBox('HERE')
                # hwb_demo = self.getNum(self.eachPerTIF[a])
                hwb_real = self.analyzeNum(self.eachPerTIF[a])
                # print('--: ', self.eachPerTIF[a])
                # print(hwb_demo)
                # hwb_real = self.analyze(hwb_demo, text)
                # print('there')
                textChecked = self.checkTXT(text)
                # print(textChecked)
                textChecked.append(self.checkforwarder(self.eachPerTIF[a]))
                # wx.MessageBox('HERE')
                b.append(s_num + str(hwb_real) + str(textChecked))
                # print(i)
                # k = str(str(n[0]['output_jpgfiles'][0]))
                if re.findall(r"\d+\s\d+k", text, re.I) != []:
                    # print(re.findall(r"\d+\s\d+k", text, re.I)[0])
                    p_w =re.findall(r"\d+\s\d+k", text, re.I)[0]
                    pack = str(re.findall(r"\d+", p_w, re.I)[0])
                    weight = str(re.findall(r"\d+", p_w, re.I)[1])
                elif  re.findall(r"\d+\s\d+.{3,4}K", text, re.I) != []:
                    p_w =  re.findall(r"\d+\s\d+.{3,4}K", text, re.I)[0]
                    pack = str(re.findall(r"\d+", p_w, re.I)[0])
                    # weight = str(re.findall(r"\d+", p_w, re.I)[1])
                    weight=''
                    if len(re.findall(r"\d+", p_w, re.I))==2:
                        weight+=str(re.findall(r"\d+", p_w, re.I)[1])
                    else:
                        weight+=str(re.findall(r"\d+", p_w, re.I)[1])
                        weight+='.'
                        weight+=str(re.findall(r"\d+", p_w, re.I)[2])

                else:
                    # print('here!')
                    pack='none'
                    weight = 'none'
                self.tiflist.append([s_num, hwb_real, textChecked,pack,weight])
                # print('hh')
                # print(self.tiflist)
                # f = open(i + ".txt", 'w', encoding='utf-8')
                # f.write(text)
                # f.close()
                # print(self.eachPerTIF[a])
                # print(hwb_demo)
                a += 1
                # print(text)
                # print(str(b))
            # for z in self.tiflist:
            #     for i in range(len(self.tiflist)-1):
            #         if z[0]==self.tiflist[i+1][0]:
            #             if z[1]!='none':
            #                 if z[2]!=[]:
            #                     pass
            #                 else:
            #                     z[2] == self.tiflist[i + 1][2]
            #             else:
            #                 z[1]==self.tiflist[i+1][1]
                item+=1
                print('TIF图片已完成 ' + str(item) + ' 单'+'\n运单号为： '+self.eachPerTIF[a-1])
                # print('tif complete: '+str(item))
                # self.process.SetLabel('tif complete: '+str(item))


        self.tifstring=str(b)
        # print(self.tifstring)
        # f=open(self.inputpath+r"\result.txt",'w',encoding='utf-8')
        # f.write(str(b))
        # f.close()
        # print('Over-------------------------------------')
        # hwb_demo = self.getNum(self.eachPerTIF[a])
        pass

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
                print('i')
                print(i)
                print('anatitle')
                print(self.anaTitle(i))
                if (os.path.splitext(child)[1] == ".PDF" or os.path.splitext(child)[1] == ".pdf" )and self.anaTitle(i):
                    # print(self.analyzeNum(i))
                    self.eachPerName.append(i)
                    self.eachPDF.append(child)
                if (os.path.splitext(child)[1] == ".PDF" or os.path.splitext(child)[1] == ".pdf" )and (not (self.anaTitle(i))) and self.contract(i):
                    # print(self.analyzeNum(i))
                    # print(i)
                    # print(child)
                    # print(dirs)
                    # print(files)
                    # print(root)
                    # print('child')
                    newi=r'\inv'+str(rename)+'.pdf'
                    newname = root + newi
                    shutil.copyfile(child, newname)
                    # os.renames(child,newname)
                    time.sleep(1)
                    self.eachperpdfinv.append(newi)
                    self.eachpdfinv.append(newname)
                    rename+=1
        print('look here------------------------------------------------')
        print(self.eachpdfinv)
        self.totalmission +=(len(self.eachPerName))

        strInputPath=''
        for i in self.eachPDF:
            strInputPath += str(i)+'\n'
        # print('length:' +str(len(self.eachPerName)))

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
            f = open(k+".txt", 'w', encoding='utf-8')
            f.write(text)
            f.close()
            hwb_demo = self.getNum(self.eachPerTIF[a])
            a += 1
            b.append(str(self.analyze(hwb_demo,text))+str(self.checkTXT(text)))


    def lie(self,event):
        starttime = datetime.datetime.now()
        # pstring=''

        threads = []
        print('here!')
        # t1=testThre.MyThread(self.pdf_ocr())
        # t2=testThre.MyThread(self.tif_ocr())
        # t1 = threading.Thread(target=self.pdf_ocr())
        # t2 = threading.Thread(target=self.tif_ocr())
        # t1.start()
        # t2.start()
        # t1.join()
        # t2.join()
        p1 = Process(target=self.pdf_ocr())
        p2 = Process(target=self.tif_ocr())
        p1.start()

        p2.start()
        p1.join()
        p2.join()









        # print('i have start')
        # threads.append(t1)
        # threads.append(t2)
        # for i in range(2):
        #     threads[i].start()
        # for i in range(2):
        #     threads[i].join()


        tifList = self.checkSupply(self.tiflist)
        pdfList = self.checkSupply(self.pdflist)
        end_list = pdfList + tifList
        # print(end_list)
        keys_ENGI = list(self.EngineerName.keys())
        a = 0
        for i in end_list:
            i_a = i[0].lower()
            i[0] = i_a
            for n in keys_ENGI:
                if i[0] == n:
                    end_list[a][2] = self.EngineerName[n]
            a += 1
        repeated_ = []
        # print('before:')
        # print(end_list)
        for i in range(len(end_list) - 1):
            if end_list[i][0] == end_list[i + 1][0]:
                # print('there is one')
                for b in range((len(end_list[i]))):
                    if end_list[i][b] == 'none':
                        end_list[i][b] = end_list[i + 1][b]
                repeated_.append(i + 1)
        # print(repeated_)
        for z in repeated_:
            # print('i will del this')
            # print(z)
            # del end_list[z]
            end_list.pop(z)
        # print('later:')
        # print(end_list)
        the_end_list = []

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
        excelname = 'import list ' + nowtitle + '.xls'
        self.write2excel(the_end_list, excelname)
        # msgDLG=wx.MessageBox('所有的数据已经识别完成并且导入到了此EXCEL表格里：'+excelname,'成功！')
        result = wx.MessageBox('所有的数据已经识别完成并且导入到了此EXCEL表格里：' + excelname + '\n当您关闭这个窗口或者点击OK时， 会自动打开此EXCEL文件~', '成功！',
                               wx.OK | wx.CANCEL | wx.ICON_EXCLAMATION)
        if result == wx.OK:
            win32api.ShellExecute(0, 'open', excelname + '.xls', '', '', 1)
            wx.Exit()
        elif result == wx.CANCEL:
            wx.Exit()
        # wx.Exit()















app = wx.App(False)
frame =  showFrame(None)
frame.Show(True)
app.MainLoop()