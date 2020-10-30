import os, re, xlrd, xlwt, pkuseg, numpy, jieba

# 字频
# 加载字频数据库
zp = xlrd.open_workbook('./corpus/CorpusCharacterlist.xls')
zp.sheet_names()
table = zp.sheet_by_name('zifreq')
charlist = table.col_values(1)
charrate = table.col_values(3)
zpDict = dict(zip(charlist, charrate))
print('字频数据库加载完成')

#词频
# 加载词频数据库
cp = xlrd.open_workbook('./corpus/CorpusWordPOSlist.xls')
cp.sheet_names()
table = cp.sheet_by_name('wordposfreq_50')
worldlist = table.col_values(1)
worldrate = table.col_values(5)
worldcate = table.col_values(3)
cpDict = dict(zip(worldlist, worldrate))
cxDict = dict(zip(worldlist, worldcate))
print('词频数据库加载完成')

filePath = './files/'
readmList = os.listdir(filePath)
print(readmList)
for fileName in readmList:
    print('开始分析'+fileName)
    # 打开阅读材料
    readm = open(filePath+fileName, 'r', encoding='UTF-8')
    # 读取内容
    readmData = readm.read()

    # 文件字频统计
    charDict = {}
    for n, char in enumerate(readmData):
        if '\u4e00'<= char<='\u9fff':
            if char in charDict:
                charDict[char] += 1
            else:
                charDict[char] = 1
        else:
            pass
    charNum1 = n
    charNum2 = len(charDict)
    
    # 字频统计结果
    charResult = []
    for char in charDict.keys():
        charRate = charDict[char]/n
        if char in zpDict.keys():
            zp = zpDict[char]
        else:
            zp = 0
        charResult.append([char, charDict[char], charRate, zp])
    totalZP = 0
    for i in range(len(charResult)):
        totalZP = totalZP + charResult[i][3]
    meanZP = totalZP/(i+1) # 平均字频

    # 文件词频统计
    wordDict1 = {}
    seg = pkuseg.pkuseg()
    biaodian = '，。、《》？、！：；“”%……（）' # 排除常用标点
    for n, word in enumerate(seg.cut(readmData)):
        if word in biaodian:
            pass
        else:
            if word in wordDict1:
                wordDict1[word] += 1
            else:
                wordDict1[word] = 1
    wordNum1 = n
    wordNum2 = len(wordDict1)
    # Num1是有重复计算数字 Num2是无重复计算
    # 词频统计结果
    wordResult1 = []
    wordResult2 = []
    for word in wordDict1.keys():
        if word in cxDict.keys():
            cx = cxDict[word]
        else:
            cx = None
        if word in cpDict.keys():
            cp = cpDict[word]
        else:
            cp = 0
            wordResult2.append([word, wordDict1[word], cp, cx])
            continue
        wordResult1.append([word, wordDict1[word], cp, cx])
    totalCP = 0
    for i in range(len(wordResult1)):
        totalCP = totalCP + wordResult1[i][2]
    meanCP = totalCP/(i+1) # 平均词频

    # 切分句子统计句长
    snts1 = re.split('，|。|？|！|……|：|;|“|”|\\n', readmData)
    snts1 = list(filter(None, snts1))
    sntsNum1 = len(snts1)
    sntlen1 = []
    sntResult1 = []
    for snt in snts1:
        sntlen1.append(len(snt))
        sntResult1.append([snt, len(snt)])
    meanSntLen1 = numpy.mean(sntlen1)

    snts2 = re.split('。|？|！|……|;|\\n', readmData)
    snts2 = list(filter(None, snts2))
    snts2 = list(filter(lambda snts2: len(snts2)!=1, snts2))
    sntsNum2 =len(snts2)
    sntlen2 = []
    sntResult2 = []
    for snt in snts2:
        sntD = re.sub('，|：|“|”|‘|’', '', snt)
        sntlen2.append(len(sntD))
        sntResult2.append([snt, len(sntD)])
    meanSntLen2 = numpy.mean(sntlen2)

    readm.close()

    # 建立结果保存excel
    excelResult = xlwt.Workbook()

    # 写入字频统计结果
    sheet1 = excelResult.add_sheet(u'字频', cell_overwrite_ok=True)
    charTitle = ['字', '字数', '占比', '字频（/百万）']
    for col in range(len(charTitle)):
        sheet1.write(0,col,charTitle[col])
    for row, item in enumerate(charResult):
        for col in range(len(item)):
            sheet1.write(row+1,col,item[col])
    print("字频统计完成")

    # 写入词频统计结果
    sheet2 = excelResult.add_sheet(u'词频', cell_overwrite_ok=True)
    wordTitle = ['词', '次数', '词频', '词性']
    for col in range(len(wordTitle)):
        sheet2.write(0,col,wordTitle[col])
    for row, item in enumerate(wordResult1):
        for col in range(len(item)):
            sheet2.write(row+1,col,item[col])
    print("词频统计完成")

    sheet2_1 = excelResult.add_sheet(u'排除词', cell_overwrite_ok=True)
    wordTitle = ['词', '次数', '词频', '词性']
    for col in range(len(wordTitle)):
        sheet2_1.write(0,col,wordTitle[col])
    for row, item in enumerate(wordResult2):
        for col in range(len(item)):
            sheet2_1.write(row+1,col,item[col])
    print("词频2统计完成")

    sheet3 = excelResult.add_sheet(u'分句-1', cell_overwrite_ok=True)
    sntTitle = ['句子', '句长']
    for col in range(len(sntTitle)):
        sheet3.write(0,col,sntTitle[col])
    for row, item in enumerate(sntResult1):
        for col in range(len(item)):
            sheet3.write(row+1,col,item[col])
    print("分句-1统计完成")

    sheet4 = excelResult.add_sheet(u'分句-2', cell_overwrite_ok=True)
    sntTitle = ['句子', '句长']
    for col in range(len(sntTitle)):
        sheet4.write(0,col,sntTitle[col])
    for row, item in enumerate(sntResult2):
        for col in range(len(item)):
            sheet4.write(row+1,col,item[col])
    print("分句-2统计完成")

    sheet5 = excelResult.add_sheet(u'统计', cell_overwrite_ok=True)
    sataTitle = ['总字数', '用字数', '平均字频', '总词数（有重复）', '总词数（无重复）', '平均词频', '平均句长-1', '平均句长-2', '最长句长']
    sataContain = [charNum1, charNum2, meanZP, wordNum1, wordNum2, meanCP, meanSntLen1, meanSntLen2, max(sntlen2)]
    for row in range(len(sataTitle)):
        sheet5.write(row,0,sataTitle[row])
        sheet5.write(row,1,sataContain[row])
    print('统计完成')
    print('分析完成：'+fileName)

    excelResult.save('./result/'+fileName+'.xls')