# -*- coding: utf-8 -*-

# Created by David Teng on 17-12-4
import codecs
import xlwt
import os
from ConfigParser import ConfigParser
from excelOpreations import excelOperations
from configobj import ConfigObj
from PIL import Image
import traceback
# import sys
# reload(sys)
# sys.setdefault
u"""
    1. 首先获取答案库数据
    2. 获取Excel表格数据
    3. 针对错题序号，从答案库中选取对应的答案填入对应的cell
    4. 针对“序号”, 检测值判断是否结束
"""


class detailsStatistics():
    u"""
        课堂详细情况统计
        目前包括课前测错题评论、南大教辅检查
    """

    def __init__(self, excelFile, startRow, startCol):
        u"""
            提供错题列的第一个cell的坐标（startRow+1 就是填写评论的列）
            :param startRow:  起始行号
            :param startCol:  起始列号
        """

        self.startRow = startRow
        self.startCol = startCol
        # self.filePath = excelFile
        # self.rb = open_workbook(excelFile)      # excel文件句柄
        # self.wr = copy(self.rb).get_sheet(0)    # excel文件中sheet句柄
        self.excelObj = excelOperations(excelFile)  # excel 文件操作对象
        pass

    def convertJpgToBmp(self, path):
        """

        :return:
        """
        tmpImg = Image.open(path)
        tmpImg.save(os.path.splitext(path)[0] + ".bmp")
        pass

    def getCommLib(self, filename):
        u"""
            获取答案库，按行读取文件，每一行对应一题的答案， 自动略过空行
            :param filename: 答案库
            :return: 返回包含答案库的dict

        """

        filename = os.path.join(os.path.abspath("."), filename)
        ansDict = {}
        filename = filename.decode('utf-8')
        with codecs.open(filename, "r", encoding='utf-8') as rd:
            lines = rd.readlines()
        index = 0
        for line in lines:
            if line.strip():
                index += 1
                ansDict[str(index)] = line.strip()
        return ansDict
        pass

    def parseTheSeriNum(self, value):
        u"""
            解析cell中的值，得到详细的序列或者'/'
            :param value:
            :return:
        """
        value = value.strip()
        # print value
        if value in ["/", ""]:
            return []
        seriList = value.split()
        ansList = []
        for num in seriList:
            if '-' in num:
                indexList = num.split('-')
                ansList.extend(range(int(indexList[0]), int(indexList[-1])+1))
            else:
                ansList.append(int(num))
        return ansList
        pass


    def isNumbers(self, value):
        """
            判断cell中的值是否是数字
            :param value:  cell值
            :return: True/False
        """
        try:
            f = float(value)
            return True
        except ValueError:
            return False
        pass

    def getlastLineNum(self):
        """
            获取表格中学生信息的最后一行的行号（这里是行号，不是行号的索引，后者等于 行号-1）
            :return: 返回行号
        """
        lineNum = self.startRow
        ans = self.excelObj.getCellValue(lineNum, 0)
        while(self.isNumbers(ans)):
            lineNum += 1
            ans = self.excelObj.getCellValue(lineNum, 0)
        return lineNum
        pass


    def getCombineAns(self, ansNumList, ansLib, splitChac="\n"):
        """

        :param ansNumList:
        :param ansLib:
        :return:
        """

        tmpstr = ""
        for i in ansNumList:
            # print ansLib[i]
            tmpstr += ansLib[str(i)]
            tmpstr += splitChac
        return tmpstr.strip()
        pass




    def fillBlanks(self, ansLib, ansCol, gradeCol, rewardWords="perfect !", finishTips="'默写'处理完成！"):
        """
            根据错题题号进行对应的知识点点拨
            :param ansLib: 知识点库
            :param ansCol: 填写点拨的列的索引
            :param gradeCol: "测试成绩" 列的索引
            :param rewardWords:  对于满分学生的夸奖
            :param finishTips: 任务完成后的提示
            :return:
        """
        ans_style = self.excelObj.setFont(u"宋体", 220,
                                          align=xlwt.Alignment.HORZ_LEFT)  # Height = 字号*20, 这里字号是11, 所以，Height=220   aliment的值设置可以追踪源码查询到
        reward_style = self.excelObj.setFont(u"宋体", 400, align=xlwt.Alignment.HORZ_CENTER,
                                             color='red')  # Height = 字号*20, 这里字号是11, 所以，Height=220   aliment的值设置可以追踪源码查询到
        # style = ans_style
        lastLineNum = self.getlastLineNum()
        # print range(self.startRow, lastLineNum)
        for lineN in range(self.startRow, lastLineNum):
            style = ans_style
            ansNum = self.excelObj.getCellValue(lineN, self.startCol)       # 获取错题值
            # print lineN, ansNum, self.startCol, self.startRow
            ansNumList = self.parseTheSeriNum(ansNum)                       # 解析错题值，输出list
            if ansNumList:                                                      # 如果list不空
                anstr = self.getCombineAns(ansNumList, ansLib)                      # 根据list中的答案序号组装结果
            elif self.isNumbers(self.excelObj.getCellValue(lineN, gradeCol)):   # 否则，如果课前测成绩一栏有数值（表示满分，没错题）
                anstr = rewardWords                                                 # 答案给“Perfect ！”
                style = reward_style
            else:                                                               # 如果成绩一栏是"/"或者空,表示缺席
                anstr = "/"                                                         # 答案给"/"
            self.excelObj.setCellValue(lineN, ansCol, anstr, style)         # self.startCol + step
        self.excelObj.saveFile()                                        # 保存文件
        print finishTips
        pass




class confObjReader():
    """
        confObj模块读取配置文件， 没有具体用到，只是留作笔记
    """

    def __init__(self, cfgFileName):
        self.fileName = cfgFileName
        self.cf = ConfigObj(self.fileName)
        pass

    def readSection(self, secName):
        # 读文件
        # value1 = self.cf['fileList']
        # value2 = self.cf['startRow']
        # value3 = self.cf['startCol']
        # print value1, value2, value3
        
        #
        section1 = self.cf[secName]
        ansDict = {}
        ansDict['ansLibName'] = section1['ansLibName']
        ansDict['fileList'] = section1['fileList']
        ansDict['startRow'] = int(section1['startRow'])
        ansDict['startCol'] = int(section1['startCol'])
        ansDict['ansCol'] = int(section1['ansCol'])
        ansDict['gradeCol'] = int(section1['gradeCol'])
        ansDict['rewardWords'] = section1['rewardWords']
        ansDict['finishTips'] = section1['finishTips']
        print ansDict
        return ansDict

        #
        # you could also write
        # value6 = self.cf['课前测']['fileList']
        # value7 = self.cf['课前测']['startCol']
        # print value6, value7
        

    def writeCfg(self):
        # 写文件
        #

        self.cf['单元测']['startRow'] = 1

        #
        # section2 = {
        #     'keyword5': value5,
        #     'keyword6': value6,
        #     'sub-section': {
        #         'keyword7': value7
        #     }
        # }
        # self.cf['section2'] = section2
        # #
        # self.cf['section3'] = {}
        # self.cf['section3']['keyword 8'] = [value8, value9, value10]
        # self.cf['section3']['keyword 9'] = [value11, value12, value13]
        #
        self.cf.write()
        pass



class confPsrseReader():

    def __init__(self, configFileName):
        self.fileName = configFileName
        self.cf = ConfigParser()
        self.cf.read(self.fileName)
        pass

    def getSecInfo(self, secName):
        """

            :param secName:
            :return:
        """

        itemList = self.cf.items(secName)
        ansDict = {}
        # ansDict['secTitle'] = secName
        for item in itemList:
            ansDict[item[0]] = item[1]
        return ansDict

        # secs = self.cf.sections()
        # print 'sections:', secs
        #
        # opts = self.cf.options("课前测")
        # print "options", opts
        #
        # kvs = self.cf.items("课前测")
        # print "课前测", kvs
        #
        # kqc_startRow = self.cf.get("课前测", "startRow")
        # kqc_startCol = self.cf.getint("课前测", "startCol")
        # print kqc_startCol, kqc_startRow

        pass

    def readConf(self):
        """
            读配置文件中的所有的信息
            :return: 返回list， 其中的item是dict类型， 一个item对应配置文件中的一个section
        """
        secList = self.cf.sections()
        ansDictList = []
        for sec in secList:
            ansDictList.append(self.getSecInfo(sec))
        return ansDictList
        pass

    def writeSections(self):

        self.cf.set("单元测", "startRow", 1)
        self.cf.write(open(self.fileName, "w"))
        pass


class goProcessing():
    """

    """

    def runPrc(self, excelName, sRow, sCol, ansLibfile, ansCol, gradeCol, rewardWrd, finiedTips):
        """
            处理表格中知识点点拨的任务
            :param excelName: 待处理的表格名称
            :param sRow:  起始行号
            :param sCol:  起始列号
            :param ansLibfile: 答案库文件名称
            :param ansCol: 点拨内容所在列
            :param gradeCol: 点拨成绩所在列
            :param rewardWrd: 表扬
            :param finiedTips: 任务完成提示
            :return: True/False
        """
        try:
            print u"开始处理文件%s ..." % excelName
            DS = detailsStatistics(excelName, sRow, sCol)
            # print 1
            testBefoClass = DS.getCommLib(ansLibfile)
            # print 2
            DS.fillBlanks(ansLib=testBefoClass, ansCol=ansCol, gradeCol=gradeCol, rewardWords=rewardWrd, finishTips=finiedTips)
            print u"完成文件%s处理...\n" % excelName
            return True
        except Exception as e:
            print e
            print traceback.format_exc()
            return False
        pass

    def getXlsfiles(self, filename):
        """
            判断是否是.xls文件
            :param filename:
            :return:
        """
        # print os.path.splitext(filename)
        if os.path.splitext(filename)[1] == ".xls":
            return True
        else:
            return False

        pass

    def getproceedfilenames(self, dirname="ready_to_process/"):
        """
            获取待处理的表格: *.xls
            :param dirname: 保存文件的目录名称（路径）
            :return: 返回包含文件的list, 或者空list
        """
        dirname = os.path.join(os.path.abspath("."), dirname)
        # dirname = dirname.replace("\\", "\\\\")
        dirname = dirname.decode('utf-8')
        print dirname, os.path.isdir(dirname)
        if os.path.isdir(dirname):
            fileList = os.listdir(dirname)
            xlsfileList = [os.path.join(dirname, f) for f in fileList if self.getXlsfiles(f)]
            return xlsfileList
        print "不存在待处理文件！"
        return []

        pass

    def main(self):
        """
            程序入口
            :return:
        """
        rdCfg = confPsrseReader("Snow_XES.cfg")
        ansDictList = rdCfg.readConf()
        for cfDict in ansDictList:
            # print isinstance(str(cfDict["fileList".lower()]), str)
            xlsfilesList = self.getproceedfilenames(str(cfDict["fileList".lower()]))
            # print xlsfilesList[0]
            for mfile in xlsfilesList:
                self.runPrc(mfile,
                            sRow=int(cfDict["startRow".lower()]),
                            sCol=int(cfDict["startCol".lower()]),
                            ansLibfile=cfDict["ansLibName".lower()],
                            ansCol=int(cfDict["ansCol".lower()]),
                            gradeCol=int(cfDict["gradeCol".lower()]),
                            rewardWrd=cfDict["rewardWords".lower()],
                            finiedTips=cfDict["finishTips".lower()]
                            )
            pass

    pass


if __name__ == '__main__':

    """
        课前测错题点拨
    """
    # excelName = "test/saturday_afternoon.xls"
    # startRow = 2
    # startCol = 3
    # DS = detailsStatistics(excelName, startRow, startCol)
    # testBefoClass = DS.getCommLib("课前测_答案库.txt")
    # DS.fillBlanks(ansLib=testBefoClass, ansCol=6, gradeCol=2, rewardWords="perfect!", finishTips="课前测处理完成！")

    """
        默写错题点拨
    """
    # excelName = "test/saturday_afternoon.xls"
    # startRow = 2
    # startCol = 16
    # WFM = DS.getCommLib("默写_答案库.txt")
    # DS1 = detailsStatistics(excelName, startRow, startCol)
    # DS1.fillBlanks(ansLib=WFM, ansCol=16, gradeCol=12, rewardWords="great!", finishTips="默写处理完成！")

    tmpobj = goProcessing()
    tmpobj.main()

    # print os.path.isdir("C:\\Users\\dell\\Documents\\PycharmEnv\\ENV\\Snow_XES\\单元测_默写".decode('utf-8'))
    # print os.listdir("C:\\Users\\dell\\Documents\\PycharmEnv\\ENV\\Snow_XES\\".decode('utf-8'))
    # path = os.path.abspath("满分.jpg")
    # tmpImg = Image.open(path)
    # print tmpImg.mode
    # tmpImg = tmpImg.convert("RGB")
    # tmpImg.save(os.path.splitext(path)[0] + ".bmp")

    pass
