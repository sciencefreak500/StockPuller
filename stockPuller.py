import csv
import os
import requests
import datetime
import threading



def createFile(name, finalList):
    filename = name + ".txt"
    file = open(filename,'w')
    for i in finalList:
        temp = ""
        for index, j in enumerate(i):
            if index == 0:
                temp = j
            else:
                temp += "," + j
        temp +="\n"
        file.write(temp)
    file.close()


def newOrder(name, testList):
    finalList = []
    for i in testList:            
        temp = [i[0],name,i[4],i[2],i[3],i[1],i[5],i[6]]
        finalList.append(temp)
    createFile(name, finalList)


def readCSV(name):
    filename = name + ".csv"
    file = open(filename, "r")
    reader = csv.reader(file)
    testList = []
    for index, i in enumerate(reader):
        if index > 0:
            testList.append(i)
    file.close()
    os.remove(filename)
    newOrder(name, testList[::-1])

    
def writeCSV(name):
    now = datetime.datetime.now()
    curr_year = str(now.year)
    curr_month = str(now.month -1)
    curr_day = str(now.day)
    past_year = str(now.year -5)
    newpath = ""
    firstpart = "http://chart.finance.yahoo.com/table.csv?s="
    newpath  = firstpart+name+"&a="+curr_month+"&b="+curr_day+"&c="+past_year+"&d="+curr_month+"&e="+curr_day+"&f="+curr_year+"&g=d&ignore=.csv"
    chartdata = requests.get(newpath)

    if "<html" in chartdata.text:
        print("Error Pulling Chart...")
    else:
        filename = name + ".csv"
        file = open(filename,'w')
        file.write(chartdata.text)
        file.close()
        readCSV(name)


def looper(f):
    for i in f:
        print(i)
        writeCSV(i.strip("\n"))

def fullList():
    bigList = []
    f = open("FullTicker.source",'r')
    for i in f:
        bigList.append(i)
    f.close()

    threadCount = 200  #EDIT NUMBER OF THREADS HERE
    threads = []
    count = 0
    div = 1/threadCount
    for i in range(threadCount):
        if count < 1:
            t = threading.Thread(target=looper, args= (bigList[round(len(bigList)*count):round(len(bigList)*(count+div))],))
            threads.append(t)
            t.start()
            count += div            
    
    
    #a = threading.Thread(target=looper, args= (bigList[:round(len(bigList)*.25)],))
    #b = threading.Thread(target=looper, args= (bigList[round(len(bigList)*.25):round(len(bigList)*.5)],))
    #c = threading.Thread(target=looper, args= (bigList[round(len(bigList)*.5):round(len(bigList)*.75)],))
    #d = threading.Thread(target=looper, args= (bigList[round(len(bigList)*.75):],))
    
    #a.start()
    #b.start()
    #c.start()
    #d.start()
    
    


fullList()
