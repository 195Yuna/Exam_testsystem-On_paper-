
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 18 09:20:50 2022

@author: yingying yu

"""

import os
import shutil
import docx
import random
import sys
import datetime


def getRandomQuestion(exam, result, testPoint, key):
    """exam :a dictionary include dif kind of tests
       result :a dic asexam
       testPoint: random dic include what we want from dif exam
       key :a number use for finally max"""
    test_paper = []
    result_paper = []
    for tPkey in testPoint.keys():
        randomList = [int(i) for i in range(len(exam[tPkey]))]
        beiXuanTi = len(exam[tPkey]) - 1
        for i in range(int(testPoint[tPkey])):
            number = random.randint(0, beiXuanTi)
            test_paper.append(exam[tPkey][randomList[number]])
            result_paper.append(result[tPkey][randomList[number]])
            randomList.remove(randomList[number])    
            beiXuanTi -=1
    return (test_paper, result_paper, key)
            
        
            
    
    
doc_list  = docx.Document('1.docx')
table = doc_list.tables
table0 = table[0]  
dictionatyList = {}
danXuan = {}
danXuan_daAn ={}
duoXuan = {}
duoXuan_daAn = {}
keyList ={}
for n in range(1,len(table0.rows)):
    rows_value = table0.rows[n]
    rows_cell = [cell.text for cell in rows_value.cells]
    if rows_cell[2] in dictionatyList.keys():
        if rows_cell[3] not in dictionatyList[rows_cell[2]]:
            dictionatyList[rows_cell[2]].append(rows_cell[3])
    else:
        dictionatyList[rows_cell[2]] =[]
        
    if rows_cell[3] not in keyList.keys():
        danXuan[rows_cell[3]] = []
        danXuan_daAn[rows_cell[3]] = []
        duoXuan[rows_cell[3]] = []
        duoXuan_daAn[rows_cell[3]] = []
        keyList[rows_cell[3]] = rows_cell[-1]
        danXuan[rows_cell[3]].append(rows_cell[6])
        danXuan_daAn[rows_cell[3]].append(rows_cell[7])
        duoXuan[rows_cell[3]].append(rows_cell[10])
        duoXuan_daAn[rows_cell[3]].append(rows_cell[11])
        
    else:
        danXuan[rows_cell[3]].append(rows_cell[6])
        danXuan_daAn[rows_cell[3]].append(rows_cell[7])
        duoXuan[rows_cell[3]].append(rows_cell[10])
        duoXuan_daAn[rows_cell[3]].append(rows_cell[11])
        


tests = {1:[],2:[]}
result = {1:[],2:[]}
isDanxuan = 0
maxdDanxuan = 6
mustBeDanxuan = 0
isduoxuan = 0
maxDuoxuan = 4
totalQuestonNumber = 0
mustBeDuoxuan = 2
randomQuestionDic ={}   #what typ queston we want how many kind and how many
randomQuestionDic[1] = {}
randomQuestionDic[2] = {}
checkKeyList = [key for key in keyList.keys()]
for i in keyList.keys():
    randNumber = random.randint(1, 2)  
    if randNumber == 1 and mustBeDuoxuan == 2:
        isDanxuan += int(keyList[i])
        if isDanxuan <= maxdDanxuan:
            randomQuestionDic[randNumber][i] = int(keyList[i])
        else:
            mustBeDuoxuan = 0
            moreFromDanxuan = isDanxuan - maxdDanxuan
            randNumber = 2
            if mustBeDanxuan == 0 and isduoxuan < maxDuoxuan:
                isduoxuan += int(keyList[i])
                randomQuestionDic[randNumber][i] = moreFromDanxuan
            
    else:
        if mustBeDanxuan == 0:
            randNumber = 2
            isduoxuan += int(keyList[i])
            if isduoxuan <= maxDuoxuan:
                randomQuestionDic[randNumber][i] = int(keyList[i])
            else:
                mustBeDanxuan = 2
                moreFromDuoXuan = isduoxuan - maxDuoxuan
                randNumber = 1
                if mustBeDuoxuan == 2 and isDanxuan < maxdDanxuan:
                    isDanxuan += int(keyList[i])
                    randomQuestionDic[randNumber][i] = moreFromDuoXuan
               

    



for key in randomQuestionDic.keys():
    if key == 1:
        questions = danXuan
        answers = danXuan_daAn
    else:
        questions = duoXuan
        answers = duoXuan_daAn
    (test, answer, typ) = getRandomQuestion(questions, answers, randomQuestionDic[key], key)
    for i in test:
        tests[int(typ)].append(i)
        result[int(typ)].append(i)
        





now = datetime.datetime.now()
otherStyleTime = now.strftime("%Y年%m月")
file_test = open("考试.txt", "w")
file_test.write("                  "+otherStyleTime +"考试\n")
number = 0
for i in tests.keys():
    for x in tests[i]:
        if int(i) == 1 and number == 0:
            file_test.writelines("i am the king"+"\n")
        else:
            if int(i) == 2 and number == len(tests[1]):
                file_test.writelines("i am the QQQ"+"\n")
        number +=1
        file_test.writelines(str(number)+". "+x+"\n")
file_test.close()


file_result = open("答案.txt", "w")
file_result.write("                  "+otherStyleTime +"答案\n")
number = 0
for i in result.keys():
   
    for x in tests[i]:
        if int(i) == 1 and number == 0:
            file_result.writelines("i am the king"+"\n")
        else:
            if int(i) == 2 and number == len(tests[1]):
                file_result.writelines("i am the QQQ"+"\n")
        number +=1
        file_result.writelines(str(number)+". "+x+"\n")
file_result.close()



