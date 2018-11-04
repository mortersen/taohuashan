import os
import re
import openpyxl
import copy
import docx
from win32com import client as wc
#定义抽取对象的数据结构
PaparInfo = {'SNo':'','Sname':'','PaparTitle':'','Tname':'',\
            'Tlevel':'','Collega':'','Major':'','Class':'',\
             'Cnabstract':'','Cnkey':'','Enabstract':'',\
             'Enkey':'','Refercontext':'','Levle':'','PaparEnTitle':''
            }
#电子表格写入列信息查询表
ExcleColumNamesTable = {'SNo':2,'Sname':3,'PaparTitle':4,'Tname':5,\
                        'Tlevel':6,'Collega':7,'Major':8,'Class':9,\
                        'Cnabstract':10,'Cnkey':11,'Enabstract':13,\
                        'Enkey':14,'Refercontext':15,'PaparEnTitle':12,'Levle':1
                        }
#获取当前文件夹下所有的文件夹路径
def getAllFolder(_pathName):
    try:
        folderlists = []
        for i,j,k in os.walk(_pathName):
            folderlists.append(i)
        return folderlists
    except  Exception as err:
        print('getAllFolder' + str(err))
#获取当前文件夹下的文件名集合
def getCurrentFilenames(folderdir):
    return os.listdir(folderdir)
#筛分离DOC和docx两种文档
def splitDocFromDocx(filenames):
    docFiles = []
    docxFiles = []
    for fn in filenames:
        if '.docx' in fn:
            docxFiles.append(fn)
        elif '.doc' in fn:
            docFiles.append(fn)
    return docFiles,docxFiles
'''
for i,j,k in myWalk:
    for ii in k:
        if ii != '':
            listFileNames.append(ii)
print(len(listFileNames))#output how many files in this flold!

for i in listFileNames:
    if '.docx' in i:
        listDocxFiles.append(i)
    elif '.doc' in i :
        listDocFiles.append(i)
print(listDocFiles)#output doc files'name
print(listDocxFiles)#output docx files's name
'''
#检查是否重复转换，获取最终转换文件名列表
def delRepeatFiles(docfiles,docxfiles):
    temp = []
    for i in docfiles:
        for j in docxfiles:
            if str(i) + 'x' == str(j):
                temp.append(i)
    for i in temp:
        docfiles.remove(i)
    return docfiles
#删除~$临时文件
def delTempDocxFiles(docxfiles):
    temp = []
    for i in docxfiles:
        if '~$' in str(i):
            temp.append(i)
    for i in temp:
        docxfiles.remove(i)
    return docxfiles
#doc2docx,返回转换列表
def doc2docx(folderdir,docfiles,docxfiles):
    if docfiles == '':
        return docxfiles
    word = wc.Dispatch("Word.Application")
    for files in docfiles:
        try:
            temp_path = folderdir + '\\' + str(files)
            temp_doc = word.Documents.Open(temp_path)
            temp_doc.SaveAs(temp_path + 'x',16)
            docxfiles.append(str(files) + 'x')
            temp_doc.Close()
        except Exception as err:
            print('doc2docx: ' + files + str(err))
    word.Quit()
    return docxfiles
'''
for i in listDocFiles:#循环将DOC2DOCX
    temp_path = pathName + '\\' + str(i)
    temp_doc = word.Documents.Open(temp_path)
    temp_doc.SaveAs(temp_path + 'x',16)
    listDocxFiles.append(str(i) + 'x')
    temp_doc.Close()
word.Quit()
'''
#删除句中英文空格
def cleanStringEnSpace(stringText):#清洗字符串半角空格
    return ''.join(stringText.split())

#删除句中全角空格
def cleanStringFullSpace(stringText):#清洗字符串全角空格
    _stringText = ''
    for uchar in stringText:
        charcode = ord(uchar)
        if charcode == 12288 :
            continue
        _stringText += chr(charcode)
    return _stringText

def getText(filename):#将DOCX转换成以自然段为单位的列表结构
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(cleanStringFullSpace(para.text.strip()))
    return fullText

#用于中文关键词，将连续空格分解为分号
def changeSpace2Splite(stringText):
    _stringText = ''
    _stringTextList = stringText.split(' ')
    for i in _stringTextList:
        if i.strip() != '':
            _stringText = _stringText  + i + '；'
    return _stringText[0:len(_stringText)-1]

'''
tempText = []#论文内容
SNo = ''#学号
Sname = ''#学生姓名
Collega = ''#学院
Major = ''#专业
Class = ''#专业
PaparTitle = ''#论文题目
Tname = ''#指导老师
Tlevel = ''#老师职称
Cnabstract = ''#中文摘要
Cnkey = ''#中文关键词
CnRefer = ''#中文参考文献
'''

#tempText = getText(r'D:\论文\121304044.docx')
#获取题目
def getPaparTitle(tempText):
    PaparTitle = ''
    for temp in tempText:
        if temp == '':
            continue
        temp = cleanStringEnSpace(temp)
        #temp = cleanStringFullSpace(temp)
        if '题目' in temp:
            PaparTitle = temp.replace('题目','')
            break
    if PaparTitle != '':
        if PaparTitle[0] == ':' or PaparTitle[0] =='：':
            PaparTitle = PaparTitle[1:]
        if PaparTitle[0] == '《' and PaparTitle[-1] == '》':
            PaparTitle = PaparTitle[1:-2]
    return PaparTitle
#获取学院和专业
def getCollegaAndMajor(tempText):
    Collega = ''
    Major = ''
    Class = ''
    for temp in tempText:
        if temp == '':
            continue
        temp = cleanStringEnSpace(temp)
        if '学院' in temp and '专业' in temp and '级' in temp:
            Class = temp.split('级')[1].replace('\t','')
            _temp = temp.split('专业')[0].replace('\t','')
            _temp = _temp.split('学院')
            Collega = _temp[0]
            Major = _temp[1]
            break
    return Collega,Major,Class

#获取学生姓名和学号
def getSnameAndSNo(tempText):
    Sname = ''
    SNo = ''
    for temp in tempText:
        if temp == '':
            continue
        temp = cleanStringEnSpace(temp)
        if '学生姓名' in temp :
            if '学号' in temp:
                Sname = temp.replace('学生姓名','')
                _temp = Sname.split('学号')
                Sname = _temp[0].replace('\t','')
                SNo = _temp[1].replace('学号','').replace('\t','')
                if Sname[0] == ':' or Sname[0] == '：':
                    Sname = Sname[1:]
                if SNo != '' and SNo[0] == ':' or SNo == '：':
                    SNo = SNo[1:]
                break
            else:
                Sname = temp.replace('学生姓名', '').replace(' ','')
                if Sname[0] == ':' or Sname[0] == '：':
                    Sname = Sname[1:]
                break
    if SNo == '':#如果学号另起一行，单独再读取一次
        for temp in tempText:
            if temp == '':
                continue
            temp = cleanStringEnSpace(temp)
            if '学号' in temp:
                SNo = temp.replace('学号','').replace(' ','')
                if SNo[0] == ':' or SNo[0] == '：':
                    SNo = SNo[1:]
                break
    return Sname,SNo
#获取指导教师
def getTname(tempText):
    Tname = ''
    for temp in tempText:
        if temp == '':
            continue
        temp = cleanStringEnSpace(temp)
        if '指导教师' in temp:
            if '职称' in temp:
                Tname = temp.replace('指导教师','')
                Tname = Tname.split('职称')[0].replace('\t','')
                break
            else:
                Tname = temp.replace('指导教师','')
                break
        elif '指导老师' in temp:
            if '职称' in temp:
                Tname = temp.replace('指导老师', '')
                Tname = Tname.split('职称')[0].replace('\t', '')
                break
            else:
                Tname = temp.replace('指导老师', '')
                break
    if Tname != '':
        if Tname[0] == ':' or Tname[0] == '：':
            Tname = Tname[1:]
    return Tname

#获取老师职称
def getTlevel(tempText):
    Tlevel = ''
    for temp in tempText:
        if temp == '':
            continue
        temp = cleanStringEnSpace(temp)
        if '职称' in temp:
            Tlevel = temp[temp.find('职称')+2:]
            break
    if Tlevel != '':
        if Tlevel[0] == '：' or Tlevel[0] == ':':
            Tlevel = Tlevel[1:]
    return Tlevel

#获取中文摘要
def getCnabstract(tempText):
    Cnabstract = ''
    for i in range(len(tempText)):
        if tempText[i] == '':
            continue
        temp = cleanStringEnSpace(tempText[i])
        #temp = cleanStringFullSpace(temp)
        if '【摘要】' in temp:
            Cnabstract = temp.split('【摘要】')[1]
            j = i + 1
            while j < len(tempText):
                if tempText[j] == '':
                    break
                temp2 = cleanStringEnSpace(tempText[j])
                #temp2 = cleanStringFullSpace(temp2)
                if '【关键词】' in temp2:
                    break
                elif '[关键词]' in temp2:
                    break
                elif '关键词：' in temp2:
                    break
                elif '关键词:' in temp2:
                    break
                elif '【关键字】' in temp2:
                    break
                elif '[关键字]' in temp2:
                    break
                elif '关键字：' in temp2:
                    break
                elif '关键字:' in temp2:
                    break
                else:
                    Cnabstract = Cnabstract + '\n' + temp2
                    j = j + 1
                    continue
            break
        if '[摘要]' in temp:
            Cnabstract = temp.split('[摘要]')[1]
            j = i + 1
            while j < len(tempText):
                if tempText[j] == '':
                    break
                temp2 = cleanStringEnSpace(tempText[j])
                #temp2 = cleanStringFullSpace(temp2)
                if '【关键词】' in temp2:
                    break
                elif '[关键词]' in temp2:
                    break
                elif '关键词：' in temp2:
                    break
                elif '关键词:' in temp2:
                    break
                elif '【关键字】' in temp2:
                    break
                elif '[关键字]' in temp2:
                    break
                elif '关键字：' in temp2:
                    break
                elif '关键字:' in temp2:
                    break
                else:
                    Cnabstract = Cnabstract + '\n' + temp2
                    j = j + 1
                    continue
            break
        if '摘要：' in temp:
            Cnabstract = temp.split('摘要：')[1]
            j = i + 1
            while j < len(tempText):
                if tempText[j] == '':
                    break
                temp2 = cleanStringEnSpace(tempText[j])
                # temp2 = cleanStringFullSpace(temp2)
                if '【关键词】' in temp2:
                    break
                elif '[关键词]' in temp2:
                    break
                elif '关键词：' in temp2:
                    break
                elif '关键词:' in temp2:
                    break
                elif '【关键字】' in temp2:
                    break
                elif '[关键字]' in temp2:
                    break
                elif '关键字：' in temp2:
                    break
                elif '关键字:' in temp2:
                    break
                else:
                    Cnabstract = Cnabstract + '\n' + temp2
                    j = j + 1
                    continue
            break
        if '摘要:' in temp:
            Cnabstract = temp.split('摘要:')[1]
            j = i + 1
            while j < len(tempText):
                if tempText[j] == '':
                    break
                temp2 = cleanStringEnSpace(tempText[j])
                # temp2 = cleanStringFullSpace(temp2)
                if '【关键词】' in temp2:
                    break
                elif '[关键词]' in temp2:
                    break
                elif '关键词：' in temp2:
                    break
                elif '关键词:' in temp2:
                    break
                elif '【关键字】' in temp2:
                    break
                elif '[关键字]' in temp2:
                    break
                elif '关键字：' in temp2:
                    break
                elif '关键字:' in temp2:
                    break
                else:
                    Cnabstract = Cnabstract + '\n' + temp2
                    j = j + 1
                    continue
            break
    if Cnabstract != '':

        if Cnabstract[0] == ':' or Cnabstract[0] == '：':
            return Cnabstract[1:]
        else:
            return Cnabstract
    return Cnabstract

#获取中文关键词
def getCnKey(tempText):
    Cnkey = ''
    i = 0
    for i in range(len(tempText)):
        temp = tempText[i].strip()
        if temp == '':
            continue

        if '【关键词】' in temp:
            temp = temp.split('【关键词】')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break
        elif '[关键词]' in temp:
            temp = temp.split('[关键词]')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break
        elif '关键词：' in temp:
            temp = temp.split('关键词：')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break
        elif '关键词:' in temp:
            temp = temp.split('关键词:')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break
        elif '【关键字】' in temp:
            temp = temp.split('【关键字】')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break
        elif '[关键字]' in temp:
            temp = temp.split('[关键字]')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break
        elif '关键字：' in temp:
            temp = temp.split('关键字：')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break
        elif '关键字:' in temp:
            temp = temp.split('关键字:')[1]
            temp2 = tempText[i + 1].strip()
            if temp2 != '' and '第一章' not in temp2 and '绪论' not in temp2 and '引言' not in temp2 and len(temp2)<10:
                temp = temp + tempText[i + 1]
            else:
                break

    if temp != '' and i != len(tempText):
        Cnkey = temp
        if ',' in Cnkey:
            Cnkey = cleanStringEnSpace(Cnkey.replace(',','；'))
        elif ';' in Cnkey:
            Cnkey = cleanStringEnSpace(Cnkey.replace(';','；'))
        elif ' ' in Cnkey:
            Cnkey = changeSpace2Splite(Cnkey)
        elif '、' in Cnkey:
            Cnkey = Cnkey.replace('、','；')
        elif '\\' in Cnkey:
            Cnkey  =  Cnkey.replace('\\','；')
    if Cnkey != '':
        if Cnkey[0] == ':' or Cnkey[0] == '：':
            return Cnkey[1:]
        else:
            return Cnkey
    return Cnkey

#获取中文参考文献和英文题目
def getCnReferContextAndEnTitle(tempText,_enTitle):
    ReferContext = ''
    EnTitle = ''
    i = 0
    for i in range(len(tempText)-1,-1,-1):
        itemText = tempText[i].strip()
        if itemText == '':
            continue
        #itemText = cleanStringFullSpace(itemText)
        if '参考文献' in itemText:
            i  = i + 1
            break
    ii = 0
    for ii in range(i,len(tempText),1):
        itemText = tempText[ii].strip()
        if itemText == '':
            continue
        #itemText == cleanStringFullSpace(itemText)
        if itemText[0] == '[' and str(itemText[1]).isdecimal() and \
                (str(itemText[2]).isdecimal()and itemText[3] == ']' or itemText[2] == ']'):
            ReferContext = itemText
            ii = ii + 1
            break
    iii = 0
    for iii in range(ii,len(tempText),1):
        itemText = tempText[iii].strip()
        #itemText == cleanStringFullSpace(itemText)
        if itemText == '':
            break
        elif itemText[0] == '[' and str(itemText[1]).isdecimal() and \
                (str(itemText[2]).isdecimal()and itemText[3] == ']' or itemText[2] == ']'):
            ReferContext = ReferContext + '；' + itemText
        else:
            ReferContext = ReferContext + itemText

    if _enTitle == '':
        iiii = 0#获取一行英文题目后退出
        for iiii in range(iii,len(tempText),1):
            itemText = tempText[iiii].strip()
            if itemText == '':
                continue
            else:
                EnTitle = itemText
                break
    else:
        EnTitle = _enTitle
    return ReferContext,EnTitle

#获取英文摘要和英文关键词
def getEnabstractAndKey(tempText):
    Enabstract = ''
    Enkey = ''
    try:
        i = 0
        for i in range(len(tempText)-1,-1,-1):#定位英文摘要位置
            itemText = tempText[i]
            if itemText.strip() == '':
                continue
            if '[Abstract]' in itemText:
                Enabstract = itemText.split('[Abstract]')[1]
                break
            elif '【Abstract】' in itemText:
                Enabstract = itemText.split('【Abstract】')[1]
                break
            elif '[abstract]' in itemText:
                Enabstract = itemText.split('[abstract]')[1]
                break
            elif '【abstract】' in itemText:
                Enabstract = itemText.split('【abstract】')[1]
                break
            elif 'abstract' in itemText:
                Enabstract = itemText.split('abstract')[1]
                break
            elif 'Abstract' in itemText:
                Enabstract = itemText.split('Abstract')[1]
                break

        if Enabstract != '':#读取摘要各段内容
            ii = 0
            for ii in range(i+1,len(tempText),1):
                itemText = tempText[ii].strip()
                if itemText == '':
                    break
                if itemText[0] == '[' or itemText[0] == '【':
                    break
                elif 'key word' in itemText or 'key Word' in itemText or 'keyword' in itemText or 'keyWord' in itemText\
                        or 'Key word' in itemText or 'Key Word' in itemText or 'Keyword' in itemText or 'KeyWord' in itemText :
                    break
                else:
                    Enabstract = Enabstract + '\n' + itemText
            j = 0
            for j in range(ii,len(tempText),1):#定位英文关键字位置
                itemText = tempText[j].strip()
                if itemText == '':
                    continue
                if itemText[0] == '[':
                    Enkey = itemText.split(']')[1].strip()
                    break
                elif itemText[0] == '【':
                    Enkey = itemText.split('】')[1].strip()
                    break
                elif 'Key words' in itemText:
                    Enkey = itemText.replace('Key words','')
                    break
                elif 'Key Words' in itemText:
                    Enkey = itemText.replace('Key Words','')
                    break
                elif 'key words' in itemText:
                    Enkey = itemText.replace('key words','')
                    break
                elif 'key Words' in itemText:
                    Enkey = itemText.replace('key Words','')
                    break
                elif 'keywords' in itemText:
                    Enkey = itemText.replace('keywords','')
                    break
                elif 'Keywords' in itemText:
                    Enkey = itemText.replace('Keywords','')
                    break
                elif 'KeyWords' in itemText:
                    Enkey = itemText.replace('KeyWords','')
                    break
                elif 'keyWords' in itemText:
                    Enkey = itemText.replace('keyWords','')
                    break
                elif 'key word' in itemText:
                    Enkey = itemText.replace('key word','')
                    break
                elif 'key Word' in itemText:
                    Enkey = itemText.replace('key Word','')
                    break
                elif 'keyword' in itemText:
                    Enkey = itemText.replace('keyword','')
                    break
                elif 'keyWord' in itemText:
                    Enkey = itemText.replace('keyWord','')
                    break
                elif 'Key word' in itemText:
                    Enkey = itemText.replace('Key word','')
                    break
                elif 'Key Word' in itemText:
                    Enkey = itemText.replace('Key Word','')
                    break
                elif 'Keyword' in itemText:
                    Enkey = itemText.replace('Keyword','')
                    break
                elif 'KeyWord' in itemText:
                    Enkey = itemText.replace('KeyWord','')
                    break
            if Enkey != '':
                if j+1 < len(tempText) :#多读一行，如果不空则为英文关键字
                    itemText = tempText[j + 1].strip()
                    if itemText != '':
                        Enkey = Enkey + itemText
                if ';'in Enkey:
                    Enkey = '；'.join(Enkey.split(';'))
                elif ','in Enkey:
                    Enkey = '；'.join(Enkey.split(','))
                elif '、'in Enkey:
                    Enkey = '；'.join(Enkey.split('、'))
                elif '.'in Enkey:
                    Enkey = '；'.join(Enkey.split('.'))

        if Enabstract != '':#除去冒号
            if Enabstract[0] == ':' or Enabstract[0] == '：':
                Enabstract = Enabstract[1:]
        if Enkey != '':
            if Enkey[0] == ':' or Enkey[0] =='：':
                Enkey = Enkey[1:]

    except Exception as err:
        print('getEnabstractAndKey' + filename + str(err))
    return Enabstract,Enkey

#判断是否含有中文
def ischinese(strtemp):
    zhPattern = re.compile(u'[\u4e00-\u9fa5]+')
    match = zhPattern.search(strtemp)
    if match:
        return True
    else: return False

#获取英专论文的中英文题目
def getEnMajorTitle(tempText):
    PaparTitle = ''
    PaparEntitle = ''
    i = 0
    for i in range(len(tempText)):
        itemText = cleanStringEnSpace(tempText[i].strip())
        if itemText == '':
            continue
        if '毕业论文' in itemText:
            for j in range(i+1,len(tempText),1):
                temp2 = tempText[j].strip()
                if temp2 == '':
                    continue
                else:
                    for k in range(j,len(tempText),1):
                        temp3 = tempText[k].strip()
                        if ischinese(temp3)== False and temp3 != '':
                            PaparEntitle = PaparEntitle + temp3
                        else:
                            break
                    for m in range(k,len(tempText),1):
                        temp4 = tempText[m].strip()
                        if temp4 == '':
                            continue
                        else:
                            PaparTitle = PaparTitle + temp4
                            if tempText[m + 1].strip() != '':
                                continue
                            else: break
                    break
            break
    return PaparTitle,PaparEntitle

'''
print(tempText)
print(getPaparTitle(tempText))
print(getSnameAndSNo(tempText))
print(getCollegaAndMajor(tempText))
print(getTname(tempText))
print(getTlevel(tempText))
print(getCnabstract(tempText))
print(getCnKey(tempText))
print(getCnReferContext(tempText))
print(getEnabstractAndKey(tempText))
'''
pathName = input('PLEASE INPUT THE PATH:\t')

listAllFolder = getAllFolder(pathName)
irow = 1
wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()


for Cdir in listAllFolder:
    Cdirname = Cdir.split('\\')[-1]
    if '一般论文' in Cdirname:
        _levle = '一般论文'
    elif '优秀论文' in Cdirname:
        _levle = '优秀论文'
    else:
        _levle = '一般论文'
    listAllFolder = []
    listFileNames = []
    listDocFiles = []
    listDocxFiles = []
    listFileNames = getCurrentFilenames(Cdir)
    listFileNames = delTempDocxFiles(listFileNames)
    listDocFiles,listDocxFiles = splitDocFromDocx(listFileNames)
    if len(listDocFiles) == 0 and len(listDocxFiles) == 0:#没有DOC和DOCX文件直接跳过
        continue
    print('正在处理提取信息的文件夹是：' + Cdir)
    listDocFiles = delRepeatFiles(listDocFiles,listDocxFiles)
    listDocxFiles = doc2docx(Cdir,listDocFiles,listDocxFiles)
    #print(len(listDocxFiles))
    for filename in listDocxFiles:
        try:
            tempText = getText(Cdir + '\\' + filename)
            #print(tempText)
            tempPaparInfo = copy.copy(PaparInfo)
            tempPaparInfo['Levle'] = _levle
            tempPaparInfo['PaparTitle'] = getPaparTitle(tempText)
            temp = getSnameAndSNo(tempText)
            tempPaparInfo['Sname'] = temp[0]
            tempPaparInfo['SNo'] = temp[1]
            if tempPaparInfo['SNo'] == '':
                tempPaparInfo['SNo'] = str(filename)
            temp = getCollegaAndMajor(tempText)
            tempPaparInfo['Collega'] = temp[0]
            tempPaparInfo['Major'] = temp[1]
            tempPaparInfo['Class'] = temp[2]
            tempPaparInfo['Tname'] = getTname(tempText)
            tempPaparInfo['Tlevel'] = getTlevel(tempText)
            tempPaparInfo['Cnabstract'] = getCnabstract(tempText)
            tempPaparInfo['Cnkey'] = getCnKey(tempText)
            temp = getCnReferContextAndEnTitle(tempText,tempPaparInfo['PaparEnTitle'])
            tempPaparInfo['Refercontext'] = temp[0]
            tempPaparInfo['PaparEnTitle'] = temp[1]
            temp = getEnabstractAndKey(tempText)
            tempPaparInfo['Enabstract'] = temp[0]
            tempPaparInfo['Enkey'] = temp[1]

            if tempPaparInfo['PaparTitle'] == '':
                temp = getEnMajorTitle(tempText)
                tempPaparInfo['PaparTitle'] = temp[0]
                tempPaparInfo['PaparEnTitle'] = temp[1]


            for ik,iv in tempPaparInfo.items():
                sheet.cell(row=irow, column=ExcleColumNamesTable.get(ik)).value = iv
            irow += 1
        #wb.save(Cdir + '\\' + Cdirname +'.xlsx' )
        except Exception as err:
            print(Cdir + filename + str(err))
            sheet.cell(row = irow,column = 2).value = str(filename)
            irow += 1
    print('             ' + Cdir + "处理完毕！")
    saveFilePath = pathName.split('\\')
    saveFileName = saveFilePath[-2] + saveFilePath[-1]
wb.save(pathName +'\\' + saveFileName + r'.xlsx')