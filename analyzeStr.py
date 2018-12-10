import re
import xlwt
import time
import os
import win32com
import win32com.client
import pyodbc
import wx
class AnalyzeStr:

    """
    这是分析字符串的文件，其中的函数被主程序导入并引用
    """
    #------------------------------------------------------------------
    with open("totalname.xlsx.txt", "r",encoding='gbk') as f:
        rawdata = []
        result_dict = {}
        # a=0
        for line in f.readlines():
            line = line.strip('\n')
            rawdata.append(line)
            sample_ = line.split(',')
            # print(sample_)
            thelist = []
            # a+=1
            #
            # print(a)
            # print(sample_)
            if len(sample_) > 2:
                for i in range(len(sample_) - 1):
                    thelist.append(sample_[i + 1])
            else:
                thelist.append(sample_[1])
            result_dict[sample_[0]] = thelist
        # print(result_dict)
    with open("supplyer.txt",'r',encoding='gbk') as z:
        supplyData=[]
        for line in z.readlines():
            line=line.strip('\n')
            supplyData.append(line)
    with open("depature.txt", 'r',encoding='gbk') as p:
        depatureData = []
        for line in p.readlines():
            line = line.strip('\n')
            depatureData.append(line)
        # print(supplyData)
            # print('result_dict:  '+str(result_dict))
    with open("forwarder.txt", 'r',encoding='gbk') as p:
        forwarderData = []
        for line in p.readlines():
            line = line.strip('\n')
            forwarderData.append(line)
    with open("term.txt", 'r',encoding='gbk') as p:
        termData = []
        for line in p.readlines():
            line = line.strip('\n')
            termData.append(line)
    #这一部分是导入分析的模板文件，这些文件都为TXT格式，存储着模板信息，在今后的
    #使用中可以向其中添加并撰写处理函数即可
    #-------------------------------------------------------------------
    seadata=('XINGANG','TIANJIN','SCHENKER')
    #seadata为处理海运时使用，有以上字符的即为海运文件
    def getNum(self,string):
        num = re.sub(r'\D', "", string)
        return num
    #-----------
    #getNum处理字符串，得到其中的数字部分
    # def analyze(self,num_,totalString):
    #     # print(num_)
    #     HWB='none'
    #     newstring = re.sub(r"\s+",'',totalString)
    #     # print(num_)
    #     # print(newstring)
    #     if num_[:-3] in newstring:
    #         HWB=num_
    #     # newnum= re.sub(num, "", totalString)
    #     # print(HWB)
    #     return HWB
    #     # print(HWB)
    #
    #---------以上部分已被删除，暂判定为无影响
    def checkTXT(self,totalString):
        # newstring = re.sub(r"\s+", '', totalString)
        newstring = totalString
        getData=[]
        # for i in self.rawdata:
        #     matchObj = re.search(i, totalString, re.M | re.I)
        #     if matchObj:
        #         getData.append(i)
        # return getData


        for i in self.result_dict:

            for x in self.result_dict[i]:
                # print('x: '+x)

                matchObj = re.search(x, newstring, re.M | re.I)
            if matchObj:
                getData.append(i)
                continue

        # print(getData)
        return getData
    #checkTXT检查分析出来的所以字符，并筛选出其中有意义的部分
    def checkSupply(self,resultStr):
        a=[]
        b=0
        supplyer='none'
        depature='none'
        forwarder='none'
        sea = 'BEIJING'
        term='none'
        # print(resultStr)
        for i in resultStr:
            a.append([i[0],i[1]])
            if i[2]!=[]:
                for k in i[2]:

                    if k in self.supplyData:
                        # print('supplyer:' + k)
                        supplyer=k
                    # if k in self.depatureData:
                    #     a[b].append(k)
                    # else:
                    #     print('no supplyer')

                    if k in self.depatureData:
                        depature=k
                    if k in self.forwarderData:
                        forwarder=k
                    if k in self.seadata:
                        sea = k
                    if k in self.termData:
                        term = k



            a[b].append(supplyer)
            a[b].append(depature)
            a[b].append(forwarder)
            a[b].append(sea)
            a[b].append(term)
            a[b].append(i[3])
            a[b].append(i[4])
            supplyer='none'
            forwarder='none'
            depature='none'
            sea = 'BEIJING'
            term = 'none'



            # if i[2]!=[]:
            #     for k in i[2]:
            #         print(k)
            #         if k in self.depatureData:
            #             a[b].append(k)
            #         else:
            #             a[b].append('none')
            # else:
            #     a[b].append('none')
            b+=1
        # print(a)

        return a
    #checkSupply检查筛选出来的部分，并将这些字符串按供应商等类别划分
    def write2excel(self,list,name):
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
        localtime = time.localtime()
        printTime=str(localtime[0])+r'/'+str(localtime[1])+r'/'+str(localtime[2])
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 4
        style = xlwt.XFStyle()
        style.pattern = pattern
        a=1

        b=0
        g=1
        header=(u'NO.','write date','Inv Date','Customs Date','ETD','ETA',u'到厂日期',u'到厂时间','MAWB','HWB','Supplier','PCS','Weight','Chargeable weight','Amount','Currency','Forwarder','Depature','Destination','Terms of Delivery','Customs No','Customs duty','Customs VAT','F freight','Cost centre','Terms of payment','Shipping Model','Buyer','PR date','Other charge','AS','MARK')
        for i in header:
            sheet.write(0,b,i)
            b+=1
        print(list)


        for i in list:
            # sheet.write(a,0,i[0])
            repeat = False
            z=0
            judge = True
            if g!=len(list):
                rawS1 = i[0]
                rawS2 = list[g][0]
                rawnum2 = int(re.sub(r'\D', '', rawS2))
                rawnum1 = int(re.sub(r'\D', '', rawS1))

                if rawnum2 - rawnum1 != 1 and rawnum2 - rawnum1 != 0:
                    judge=False
                    gap = rawnum2 - rawnum1
                    for z in range(gap - 1):
                        # print(a + z + 1)
                        # print('s' + str(rawnum1 + z + 1))
                        sheet.write(a + z + 1, 0, 's' + str(rawnum1 + z + 1))
                        sheet.write(a + z + 1, 1, printTime)
                elif rawnum2 - rawnum1 == 0:
                    # print('I am repeated')
                    repeat=True

                else:
                    judge =True
            if repeat==False:
                sheet.write(a, 1, printTime)
                sheet.write(a, 9, i[1])
                sheet.write(a, 10, i[2])
                sheet.write(a, 17, i[3])
                sheet.write(a, 16, i[4])
                # sheet.write(a,18,'BEIJING')
                sheet.write(a, 17 + 8, '90')
                if i[4] == 'HIASIANG' or i[5]=='TIANJIN' or i[5]=='XINGANG':
                    sheet.write(a, 17 + 9, 'SEA')
                    sheet.write(a, 0, i[0], style)
                    sheet.write(a, 18, 'TIANJIN')
                else:
                    sheet.write(a, 17 + 9, 'AIR')
                    sheet.write(a, 0, i[0])
                    sheet.write(a, 18, 'BEIJING')
                # sheet.write(a,17+9,'AIR')
                sheet.write(a, 17 + 14, 'By COMPUTER')
            else:
                a=a-1
            g += 1
            if judge == True:
                a+=1
            else:
                a += (z + 2)



        book.save(name+".xls")
    #write2excel将结果写进Excel中，虽仍在使用中，但已不再维护，结果请全部参考写入数据库中的部分
    def analyzeNum(self,rawstring):
        dsv_determin = self.dsv_(rawstring)
        expeditor_determine = self.expeditor_(rawstring)
        kwe_determine = self.kwe_(rawstring)
        hhe_determine = self.hhe_(rawstring)
        ups_determine = self.ups_(rawstring)
        dhl_determine = self.dhl_(rawstring)
        fedex_determine = self.fedex_(rawstring)
        hiasiang_determine = self.hiasiang_(rawstring)
        agility_determine = self.agility_(rawstring)
        mlg_determine = self.mlg_(rawstring)
        dgf_determine = self.dgf_(rawstring)
        schenker_determine = self.schenker_(rawstring)
        soonest_determine=self.soonest_(rawstring)
        # if dsv_determin!='none':
        #     print(dsv_determin)

        # if re.findall(r"prg\d+",rawstring, re.I)!=[]:
        #     return re.findall(r"prg\d+",rawstring, re.I)[0]
        if dsv_determin!='none':
            return dsv_determin
        # elif re.findall(r"LOCL",rawstring, re.I)!=[]:
        #     return re.findall(r"^\d+",rawstring, re.I)[0]
        elif expeditor_determine!='none':
            return expeditor_determine
        elif soonest_determine!='none':
            return soonest_determine
        elif schenker_determine!='none':
            return schenker_determine
        elif mlg_determine!='none':
            return mlg_determine
        elif dgf_determine!= 'none':
            return dgf_determine
        elif agility_determine!='none':
            return agility_determine
        elif kwe_determine!='none':
            return kwe_determine
        elif hhe_determine!='none':
            return hhe_determine
        elif ups_determine!='none':
            return ups_determine
        elif dhl_determine!='none':
            # print('dhl')
            return dhl_determine
        elif hiasiang_determine!='none':
            return hiasiang_determine
        elif fedex_determine!='none':
            return fedex_determine
        # elif re.findall(r"ZRH\d+",rawstring, re.I)!=[]:
        #     return re.findall(r"ZRH\d+",rawstring)[0]
        # elif re.findall(r"HKG\d+",rawstring, re.I)!=[]:
        #     return re.findall(r"HKG\d+",rawstring)[0]
        # elif re.findall(r"NUE\d+",rawstring, re.I)!=[]:
        #     return re.findall(r"NUE\d+",rawstring)[0]
        # elif re.findall(r"7[a-z][a-z]\d+",rawstring, re.I)!=[]:
        #     return re.findall(r"7[a-z][a-z]\d+",rawstring, re.I)[0]
        # elif self.dsv_(rawstring)!='none':
        #     return
        # elif re.findall(r"^\d+",rawstring, re.I)!=[]:
        #     return re.findall(r"^\d+",rawstring, re.I)[0]
        else:
            return 'none'
    #analyzeNum按找给与的正则表达式分析每一单的文件名，如果正确就返回运单号
    def anaTitle(self,rawi):
        dsv_determin = self.dsv_(rawi)
        expeditor_determine = self.expeditor_(rawi)
        kwe_determine = self.kwe_(rawi)
        hhe_determine = self.hhe_(rawi)
        ups_determine = self.ups_(rawi)
        dhl_determine = self.dhl_(rawi)
        fedex_determine = self.fedex_(rawi)
        hiasiang_determine = self.hiasiang_(rawi)
        agility_determine = self.agility_(rawi)
        mlg_determine = self.mlg_(rawi)
        dgf_determine = self.dgf_(rawi)
        schenker_determine=self.schenker_(rawi)
        soonest_determine=self.soonest_(rawi)
        # inv_determine = self.check_inv(rawi)
        # contract_determine = self.check_contract(rawi)
        #--------------------------------------------------
        # if re.findall(r"prg\d+",rawi, re.I)!=[]:
        #     return True
        # elif re.findall(r"^450\d+",rawi, re.I)!=[]:
        #     return False
        # elif re.findall(r"^ISTCH\d+",rawi, re.I)!=[]:
        #     return False
        # elif re.findall(r"Inv",rawi, re.I)!=[]:
        #     return False
        # elif re.findall(r"inv",rawi, re.I)!=[]:
        #     return False
        # elif re.findall(r"LOCL",rawi, re.I)!=[]:
        #     return True
        # elif re.findall(r"AWB",rawi, re.I)!=[]:
        #     return True
        # elif re.findall(r"7[a-z][a-z]\d+",rawi, re.I)!=[]:
        #     return True
        # elif ( re.findall(r"^\d+",rawi, re.I )!=[] and len(re.findall(r"^\d+",rawi, re.I)[0])>5 and len(re.findall(r"^\d+",rawi, re.I)[0])<13):
        #     # print('this--')
        #     return True
        #------------------------------------
        if dsv_determin!='none':
            return dsv_determin
        # elif re.findall(r"LOCL",rawstring, re.I)!=[]:
        #     return re.findall(r"^\d+",rawstring, re.I)[0]
        # elif contract_determine=='none':
        #     return True
        # elif inv_determine =='none':
        #     return True
        elif schenker_determine!='none':
            return True
        elif soonest_determine!='none':
            return True
        elif expeditor_determine!='none':
            return True
        elif mlg_determine!='none':
            return True
        elif dgf_determine!='none':
            return True
        elif agility_determine!='none':
            return True
        elif kwe_determine!='none':
            return True
        elif hhe_determine!='none':
            return True
        elif ups_determine!='none':
            return True
        elif dhl_determine!='none':
            return True
        elif hiasiang_determine!='none':
            return True
        elif fedex_determine!='none':
            return True
        # elif re.findall(r"^\d+",rawi, re.I)!=[]:
        #     return True

        else:
            return False
    #anaTitle分析文件名，如果符合就返回真。
    def checkforwarder(self,rawi):
        dsv_determin = self.dsv_(rawi)
        expeditor_determine = self.expeditor_(rawi)
        kwe_determine = self.kwe_(rawi)
        hhe_determine = self.hhe_(rawi)
        ups_determine = self.ups_(rawi)
        dhl_determine = self.dhl_(rawi)
        fedex_determine = self.fedex_(rawi)
        hiasiang_determine = self.hiasiang_(rawi)
        agility_determine = self.agility_(rawi)
        mlg_determine = self.mlg_(rawi)
        dgf_determine = self.dgf_(rawi)
        schenker_determine=self.schenker_(rawi)
        soonest_determine=self.soonest_(rawi)
        if dsv_determin!='none':
            return 'DSV'
        # elif re.findall(r"LOCL",rawstring, re.I)!=[]:
        #     return re.findall(r"^\d+",rawstring, re.I)[0]
        elif schenker_determine!='none':
            return 'SCHENKER'
        elif soonest_determine!='none':
            return 'SOONEST'
        elif expeditor_determine!='none':
            return 'EXPEDITOR'
        elif mlg_determine!='none':
            return 'MLG'
        elif dgf_determine!='none':
            return 'DGF'
        elif kwe_determine!='none':
            return 'KWE'
        elif agility_determine!='none':
            return 'AGILITY'
        elif hhe_determine!='none':
            return 'HHE'
        elif ups_determine!='none':
            return 'UPS'
        elif dhl_determine!='none':
            return 'DHL'
        elif hiasiang_determine!='none':
            return 'HIASIANG'
        elif fedex_determine!='none':
            return 'FEDEX'
        # elif re.findall(r"^\d+",rawi, re.I)!=[]:
        #     return True

        else:
            return 'none'
    #checkforwarder分析文件名，返回其运输商
    #---------------------------------------------
    def dsv_(self,test):
        if re.findall(r"prg\d+",test, re.I)!=[]:
            return re.findall(r"prg\d+",test, re.I)[0]
        elif re.findall(r"ZRH\d+",test, re.I)!=[]:
            return re.findall(r"ZRH\d+",test)[0]
        elif re.findall(r"HKG\d+",test, re.I)!=[]:
            return re.findall(r"HKG\d+",test)[0]
        elif re.findall(r"NUE\d+",test, re.I)!=[]:
            return re.findall(r"NUE\d+",test)[0]
        elif re.findall(r"POZ\d+",test, re.I)!=[]:
            return re.findall(r"POZ\d+",test)[0]
        elif re.findall(r"BOM\d+",test, re.I)!=[]:
            return re.findall(r"BOM\d+",test)[0]
        elif re.findall(r"SIN\d+",test, re.I)!=[]:
            return re.findall(r"SIN\d+",test)[0]
        elif re.findall(r"STO\d+",test, re.I)!=[]:
            return re.findall(r"STO\d+",test)[0]
        elif re.findall(r"ATL\d+",test, re.I)!=[]:
            return re.findall(r"ATL\d+",test)[0]
        elif re.findall(r"PDA\d+",test, re.I)!=[]:
            return re.findall(r"PDA\d+",test)[0]
        else:
            return 'none'
    def expeditor_(self,test):
        if re.findall(r"^4913\d+",test, re.I)!=[]:
            return re.findall(r"^4913\d+",test, re.I)[0]
        else:
            return 'none'
    def kwe_(self,test):
        if re.findall(r"^5800\d+",test, re.I)!=[]:
            return re.findall(r"^5800\d+",test, re.I)[0]
        elif re.findall(r"LOCL", test, re.I) != []:
            return re.findall(r"^\d+", test, re.I)[0]
        else:
            return 'none'
    def hhe_(self,test):
        if re.findall(r"hhe-\d+.\d+",test, re.I)!=[]:
            return re.findall(r"hhe-\d+.\d+",test, re.I)[0]
        else:
            return 'none'
    def ups_(self,test):
        if re.findall(r"1z\S{16}",test, re.I)!=[]:
            return re.findall(r"1z\S{16}",test, re.I)[0]
        else:
            return 'none'
    def dhl_(self,test):
        # if re.findall(r"7[a-z][a-z]\d+", test, re.I) != []:
        #     return re.findall(r"7[a-z][a-z]\d+", test, re.I)[0]
        check=False
        # print(test)

        if '.tif' in test:
            new_= test.split('.tif')[0]
            # print(new_)
            if len(new_)==10:
                check=True
        # print(check)

        if (re.findall(r"^\d{10}", test, re.I) != []) and (re.findall(r"awb", test, re.I) != []) or (check==True) or (re.findall(r"^\d{10}[.]pdf", test, re.I) != []):
            return re.findall(r"^\d{10}", test, re.I)[0]
        else:
            return 'none'
    def dgf_(self,test):
        if re.findall(r"7[a-z][a-z]\d{4}", test, re.I) != []:
             return re.findall(r"7[a-z][a-z]\d{4}", test, re.I)[0]
        else:
            return 'none'
    def fedex_(self,test):
        # check = False
        # print('test',test)
        #
        # if '.pdf' in test:
        #     new_ = test.split('.pdf')[0]
        #     print(new_)
        #     if len(new_) == 12:
        #         check = True
        # print(check)
        if re.findall(r"^\d{12}\D+", test, re.I) != [] :
            return re.findall(r"^\d{12}", test, re.I)[0]
        else:
            return 'none'
    def hiasiang_(self,test):
        if re.findall(r"^7721\d{7}", test, re.I) != []:
            return re.findall(r"^7721\d+", test, re.I)[0]
        else:
            return 'none'
    def agility_(self,test):
        if (re.findall(r"TH\d{8}", test, re.I) != []):
            return re.findall(r"TH\d{8}", test, re.I)[0]
        else:
            return 'none'
    def mlg_(self,test):
        if (re.findall(r"MLG.\d{8}", test, re.I) != []):
            return re.findall(r"MLG.\d{8}", test, re.I)[0]
        else:
            return 'none'
    def schenker_(self,test):
        print(test)
        if (re.findall(r"USCHI\d{10}", test, re.I) != []):
            return re.findall(r"USCHI\d{10}", test, re.I)[0]
        else:
            return 'none'
    def soonest_(self,test):
        if (re.findall(r"SEC\d{8}", test, re.I) != []):
            return re.findall(r"SEC\d{8}", test, re.I)[0]
        else:
            return 'none'
    #以上部分为每家运输商的正则表达式，用以分析文件名以得出运单号
    #-----------------------------------

    def contract(self,test):
        if (re.findall(r"^(450\d{7})", test, re.I) != []):
            return False
        else:
            return True
    #contract分析合同号以避开此类文件
    def accdb_(self,list,DBfile):
        # conn = win32com.client.Dispatch(r"ADODB.Connection")
        # DSN = 'PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE =大表录入程序.mdb'
        # conn.Open(DSN)
        # rs = win32com.client.Dispatch(r'ADODB.Recordset')
        # rs_name = 'importList'
        # rs.Open('[' + rs_name + ']', conn, 1, 3)

        #-----------------------------------------------
        #选择数据库--
        # wildcard = u"Access 数据库 (*.accdb)|*.accdb" \
        #
        # dlg = wx.FileDialog(self, message=u"选择文件",
        #                     defaultDir=os.getcwd(),
        #                     defaultFile="",
        #                     wildcard=wildcard,
        #                     style=wx.OPEN  | wx.CHANGE_DIR)
        # if dlg.ShowModal() == wx.ID_OK:
        #     DBfile = dlg.GetPaths()
        # dlg.Destroy()

        #--------
        # DBfile = r'e:\packingOCR\大表数据库.accdb'
        conn = pyodbc.connect(
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + DBfile + ";Uid=;Pwd=;charset='utf-8';")
        cursor = conn.cursor()
        rs_name = 'importList'
        #--------------------------------------
        localtime = time.localtime()
        printTime = str(localtime[0]) + r'/' + str(localtime[1]) + r'/' + str(localtime[2])
        a = 1
        b = 0
        g = 1
        # for
        each_first={}
        posi=0
        for each in list:
            each_first[each[0]]=posi
            posi+=1
        print(each_first)
        for raw_title in self.total_dir:

            if raw_title in each_first.keys():
                i=list[each_first[raw_title]]
                sentence3 = "VALUES ('" + raw_title + "','" + printTime + "','" + i[
                    1] + "','" + i[2] + "','" + i[4] + "','" + i[3] + "','" + i[6] + "','" + i[7] + "','" + i[
                                8] + "','" + i[8]
                if i[4] == 'HIASIANG' or i[5] == 'TIANJIN' or i[5] == 'XINGANG' or i[4]=='SCHENKER':
                    # sheet.write(a, 17 + 9, 'SEA')
                    # sheet.write(a, 0, i[0], style)
                    # sheet.write(a, 18, 'TIANJIN')
                    sentence3 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
                else:
                    # sheet.write(a, 17 + 9, 'AIR')
                    # sheet.write(a, 0, i[0])
                    # sheet.write(a, 18, 'BEIJING')
                    sentence3 += "','" + "BEIJING" + "','" + "90" + "','" + "AIR" + "','" + "BY_PC" + "')"
                sql3 = "Insert Into " + rs_name + "([No],[writeDate],[HWB],[Supplier],[Forwarder],[Departure],[TermsOfDelivery],[PCS],[Weight],[ChargeableWeight],[Destination],[TermsOfPayment],[ShippingModel],[MARK])" + sentence3
                cursor.execute(sql3)
                conn.commit()




            else:
                print('?----------'+raw_title)
                sentence1 = "VALUES ('" + raw_title + "','" + printTime + "')"
                sql1 = "Insert Into " + rs_name + "([No],[writeDate])" + sentence1
                print(sql1)
                conn.execute(sql1)
                conn.commit()
        print('Already Complete, Changes are made in your Access.')

        cursor.close()
        conn.close()



            #         sentence4 = "VALUES ('" + 's' + str(rawnum1) + "','" + printTime + "','" + i[
            #             1] + "','" + i[2] + "','" + i[4] + "','" + i[3]+ "','" + i[6]+ "','" + i[7]+ "','" + i[8]+ "','" + i[8]
            #         if i[4] == 'HIASIANG' or i[5] == 'TIANJIN' or i[5] == 'XINGANG':
            #             # sheet.write(a, 17 + 9, 'SEA')
            #             # sheet.write(a, 0, i[0], style)
            #             # sheet.write(a, 18, 'TIANJIN')
            #             sentence4 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
            #         else:
            #             # sheet.write(a, 17 + 9, 'AIR')
            #             # sheet.write(a, 0, i[0])
            #             # sheet.write(a, 18, 'BEIJING')
            #             sentence4 += "','" + "BEIJING" + "','" + "90" + "','" + "AIR" + "','" + "BY_PC" + "')"
            #         sql4 = "Insert Into " + rs_name + "([No],[writeDate],[HWB],[Supplier],[Forwarder],[Departure],[TermsOfDelivery],[PCS],[Weight],[ChargeableWeight],[Destination],[TermsOfPayment],[ShippingModel],[MARK])" + sentence4
            #         cursor.execute(sql4)
            #         conn.commit()
            # else:
            #     # print(i)
            #     rawS1 = i[0]
            #     rawnum1 = int(re.sub(r'\D', '', rawS1))
            #     sentence5 = "VALUES ('" + 's' + str(rawnum1) + "','" + printTime + "','" + i[
            #         1] + "','" + i[2] + "','" + i[4] + "','" + i[3] + "','" + i[6]+ "','" + i[7]+ "','" + i[8]+ "','" + i[8]
            #     if i[4] == 'HIASIANG' or i[5] == 'TIANJIN' or i[5] == 'XINGANG':
            #         # sheet.write(a, 17 + 9, 'SEA')
            #         # sheet.write(a, 0, i[0], style)
            #         # sheet.write(a, 18, 'TIANJIN')
            #         sentence5 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
            #     else:
            #         # sheet.write(a, 17 + 9, 'AIR')
            #         # sheet.write(a, 0, i[0])
            #         # sheet.write(a, 18, 'BEIJING')
            #         sentence5 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
            #     sql5 = "Insert Into " + rs_name + "([No],[writeDate],[HWB],[Supplier],[Forwarder],[Departure],[TermsOfDelivery],[PCS],[Weight],[ChargeableWeight],[Destination],[TermsOfPayment],[ShippingModel],[MARK])" + sentence5
            #     cursor.execute(sql5)
            #     conn.commit()

            # if repeat==False:
            #     # sheet.write(a, 1, printTime)
            #     # sheet.write(a, 9, i[1])
            #     # sheet.write(a, 10, i[2])
            #     # sheet.write(a, 17, i[3])
            #     # sheet.write(a, 16, i[4])
            #     # # sheet.write(a,18,'BEIJING')
            #     # sheet.write(a, 17 + 8, '90')
            #     sentence2="VALUES ('" + 's' + str(rawnum1 + z + 1) + "','" + printTime + "','" + i[1] + "','" + i[2] + "','" + i[4] + "','" + i[3]
            #     if i[4] == 'HIASIANG':
            #         # sheet.write(a, 17 + 9, 'SEA')
            #         # sheet.write(a, 0, i[0], style)
            #         # sheet.write(a, 18, 'TIANJIN')
            #         sentence2 += "','"+"TIANJIN"+"','"+"90"+"','"+"SEA"+"','"+"BY_PC"+"')"
            #     else:
            #         # sheet.write(a, 17 + 9, 'AIR')
            #         # sheet.write(a, 0, i[0])
            #         # sheet.write(a, 18, 'BEIJING')
            #         sentence2 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
            #     # sheet.write(a,17+9,'AIR')
            #     # sheet.write(a, 17 + 14, 'By COMPUTER')
            #     sql2 = "Insert Into " + rs_name+ "([No],[writeDate],[HWB],[Supplier],[Forwarder],[Departure],[Destination],[TermsOfPayment],[ShippingModel],[MARK])" + sentence2
            #     conn.Execute(sql2)
            # else:
            #     a=a-1
            # g += 1
            # if judge == True:
            #     a+=1
            # else:
            #     a += (z + 2)

    #accdb_向数据库里写入并更新数据


    # def adb_(self,list):
    #     conn = win32com.client.Dispatch(r"ADODB.Connection")
    #     DSN = 'PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE =大表录入程序.mdb'
    #     conn.Open(DSN)
    #     rs = win32com.client.Dispatch(r'ADODB.Recordset')
    #     rs_name = 'importList'
    #     rs.Open('[' + rs_name + ']', conn, 1, 3)
    #
    #
    #
    #     localtime = time.localtime()
    #     printTime = str(localtime[0]) + r'/' + str(localtime[1]) + r'/' + str(localtime[2])
    #     a = 1
    #     b = 0
    #     g = 1
    #     for i in list:
    #         repeat = False
    #         z = 0
    #         judge = True
    #         if g != len(list):
    #             rawS1 = i[0]
    #             rawS2 = list[g][0]
    #             rawnum2 = int(re.sub(r'\D', '', rawS2))
    #             rawnum1 = int(re.sub(r'\D', '', rawS1))
    #             if rawnum2 - rawnum1 != 1 and rawnum2 - rawnum1 != 0:
    #                 judge=False
    #                 gap = rawnum2 - rawnum1
    #                 sentence3 = "VALUES ('" + 's' + str(rawnum1 ) + "','" + printTime + "','" + i[
    #                     1] + "','" + i[2] + "','" + i[4] + "','" + i[3]
    #                 if i[4] == 'HIASIANG' or i[5]=='TIANJIN' or i[5]=='XINGANG' :
    #                     # sheet.write(a, 17 + 9, 'SEA')
    #                     # sheet.write(a, 0, i[0], style)
    #                     # sheet.write(a, 18, 'TIANJIN')
    #                     sentence3 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
    #                 else:
    #                     # sheet.write(a, 17 + 9, 'AIR')
    #                     # sheet.write(a, 0, i[0])
    #                     # sheet.write(a, 18, 'BEIJING')
    #                     sentence3 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
    #                 sql3 = "Insert Into " + rs_name + "([No],[writeDate],[HWB],[Supplier],[Forwarder],[Departure],[Destination],[TermsOfPayment],[ShippingModel],[MARK])" + sentence3
    #                 conn.Execute(sql3)
    #                 for z in range(gap - 1):
    #                     # print(a + z + 1)
    #                     # print('s' + str(rawnum1 + z + 1))
    #                     # sheet.write(a + z + 1, 0, 's' + str(rawnum1 + z + 1))
    #                     # sheet.write(a + z + 1, 1, printTime)
    #                     sentence2 = "VALUES ('" + 's' + str(rawnum1 + z + 1) + "','" + printTime + "','" + i[
    #                         1] + "','" + i[2] + "','" + i[4] + "','" + i[3]
    #
    #                     # print('1')
    #                     sentence1 = "VALUES ('" + 's' + str(rawnum1 + z + 1) + "','" + printTime +"')"
    #                     sql1="Insert Into " + rs_name + "([No],[writeDate])" + sentence1
    #                     conn.Execute(sql1)
    #
    #             else:
    #                 sentence4 = "VALUES ('" + 's' + str(rawnum1 ) + "','" + printTime + "','" + i[
    #                     1] + "','" + i[2] + "','" + i[4] + "','" + i[3]
    #                 if i[4] == 'HIASIANG'or i[5]=='TIANJIN' or i[5]=='XINGANG':
    #                     # sheet.write(a, 17 + 9, 'SEA')
    #                     # sheet.write(a, 0, i[0], style)
    #                     # sheet.write(a, 18, 'TIANJIN')
    #                     sentence4 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
    #                 else:
    #                     # sheet.write(a, 17 + 9, 'AIR')
    #                     # sheet.write(a, 0, i[0])
    #                     # sheet.write(a, 18, 'BEIJING')
    #                     sentence4 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
    #                 sql4 = "Insert Into " + rs_name + "([No],[writeDate],[HWB],[Supplier],[Forwarder],[Departure],[Destination],[TermsOfPayment],[ShippingModel],[MARK])" + sentence4
    #                 conn.Execute(sql4)
    #
    #         # if repeat==False:
    #         #     # sheet.write(a, 1, printTime)
    #         #     # sheet.write(a, 9, i[1])
    #         #     # sheet.write(a, 10, i[2])
    #         #     # sheet.write(a, 17, i[3])
    #         #     # sheet.write(a, 16, i[4])
    #         #     # # sheet.write(a,18,'BEIJING')
    #         #     # sheet.write(a, 17 + 8, '90')
    #         #     sentence2="VALUES ('" + 's' + str(rawnum1 + z + 1) + "','" + printTime + "','" + i[1] + "','" + i[2] + "','" + i[4] + "','" + i[3]
    #         #     if i[4] == 'HIASIANG':
    #         #         # sheet.write(a, 17 + 9, 'SEA')
    #         #         # sheet.write(a, 0, i[0], style)
    #         #         # sheet.write(a, 18, 'TIANJIN')
    #         #         sentence2 += "','"+"TIANJIN"+"','"+"90"+"','"+"SEA"+"','"+"BY_PC"+"')"
    #         #     else:
    #         #         # sheet.write(a, 17 + 9, 'AIR')
    #         #         # sheet.write(a, 0, i[0])
    #         #         # sheet.write(a, 18, 'BEIJING')
    #         #         sentence2 += "','" + "TIANJIN" + "','" + "90" + "','" + "SEA" + "','" + "BY_PC" + "')"
    #         #     # sheet.write(a,17+9,'AIR')
    #         #     # sheet.write(a, 17 + 14, 'By COMPUTER')
    #         #     sql2 = "Insert Into " + rs_name+ "([No],[writeDate],[HWB],[Supplier],[Forwarder],[Departure],[Destination],[TermsOfPayment],[ShippingModel],[MARK])" + sentence2
    #         #     conn.Execute(sql2)
    #         # else:
    #         #     a=a-1
    #         # g += 1
    #         # if judge == True:
    #         #     a+=1
    #         # else:
    #         #     a += (z + 2)
    #     print('Already Complete')
    #
    #
    #
    #
    #
    #
    #     conn.Close()
    # if __name__ == '__main__':
    #     a='hhh updated'
    #     def contract(test):
    #         if (re.findall(r"^(450\d{7})", test, re.I) != []):
    #             return False
    #         else:
    #             return True
    #     print(contract(a))
# def check_inv(self,test):
    #     if re.findall(r"inv", test, re.I) != []:
    #         return 'none'
    #     else:
    #         return 'nice'
    # def check_contract(self,test):
    #     if re.findall(r"^450", test, re.I) != []:
    #         return 'none'
    #     else:
    #         return 'nice'
    # print(self.anaTitle('4506011173.PDF'))










