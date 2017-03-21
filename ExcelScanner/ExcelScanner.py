from tkinter import *
from tkinter import filedialog as fd
import shutil
import os
import threading
import xlrd
import xlwt
import time
from xlutils.copy import copy
import pandas as pd
from os import listdir
from os.path import isfile, join
import win32com

import win32com.client as win32
import pythoncom

txtLoc = 'C:\GM1000k'
threadCount = 20

threads = []
zero = []

verifylist = [f for f in listdir(txtLoc) if isfile(join(txtLoc, f))]

def run_in_thread(xl_id):
    # Initialize
    pythoncom.CoInitialize()

    # Get instance from the id
    xl = win32com.client.Dispatch(
            pythoncom.CoGetInterfaceAndReleaseStream(xl_id, pythoncom.IID_IDispatch)
    )
    time.sleep(5)

    

def populateChart(chname, sheet, book):
    chartdata = []
    
    with open(join(txtLoc,chname),'r') as f:
        for line in f:
            chartdata.append(line.strip('\n').split(','))
    addon = 1800 - len(chartdata)
    for i in range(addon):
        chartdata.append([" "," "," "," "," "," "," "," "," "])


    sheet.Range(sheet.Cells(1,1),sheet.Cells(len(chartdata),len(chartdata[0]))).Value = chartdata            
    #print(sheet.Range("GB2:HW2").Value)
    b = [list(x) for x in sheet.Range("GB2:HW2").Value]
    zero.append(b[0])
    verifylist.remove(chname)
    book.Save()
    
    

def looper(f,excelfile):

    pythoncom.CoInitialize()

    excelApp = win32com.client.DispatchEx("Excel.Application")

    myStream = pythoncom.CreateStreamOnHGlobal()    
    pythoncom.CoMarshalInterface(myStream, 
                                 pythoncom.IID_IDispatch, 
                                 excelApp._oleobj_, 
                                 pythoncom.MSHCTX_LOCAL, 
                                 pythoncom.MSHLFLAGS_TABLESTRONG)    

    excelApp = None

    myStream.Seek(0,0)
    myUnmarshaledInterface = pythoncom.CoUnmarshalInterface(myStream, pythoncom.IID_IDispatch)    
    excel = win32com.client.Dispatch(myUnmarshaledInterface)
   
    #excel = win32.gencache.EnsureDispatch('Excel.Application')
    print('init Looping')

    indiv = excel.Workbooks.Open(join(os.getcwd(),excelfile))
    indivsheet = indiv.Worksheets("data")
    for i in f:
        print(i)
        populateChart(i, indivsheet , indiv)

    indiv.Close()
    excel.Quit()
    
    # Clear the stream now that we have finished
    myStream.Seek(0,0)
    pythoncom.CoReleaseMarshalData(myStream)

    myUnmarshaledInterface = None
    excel = None
    myStream = None

    pythoncom.CoUninitialize()

def getfilePath():
    root = Tk()
    root.withdraw()
    root.update()
    file = fd.askopenfilename()
    if file: 
        print(file)
    root.destroy()
    return file

def leanmean(srcpath):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    #shutil.copy2(srcpath,join(os.getcwd(),'lean_sourcefile.xls'))
    fn = join(os.getcwd(),'lean_sourcefile.xls')
    os.remove(fn) if os.path.exists(fn) else None
    print("Making Lean File...")
    orig = excel.Workbooks.Open(srcpath)
    origsheet = orig.Worksheets("data")
    newbook = excel.Workbooks.Add()
    newsheet = newbook.Worksheets.Add()
    origsheet.Copy(newsheet)
    excel.DisplayAlerts = False
    newbook.SaveAs(join(os.getcwd(),'lean_sourcefile.xls'))
    excel.DisplayAlerts = True
    newbook.Close()
    orig.Close()
    excel.Application.Quit()

    print('Lean File Created')

def writeZero():
    print("verify",verifylist, len(verifylist))
    if len(verifylist) > 0:
        looper(verifylist,'leanmean'+ str(0)+'.xls')
        writeZero()
    print("write to zero file")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    zerobook = excel.Workbooks.Add()
    zerosheet = zerobook.Worksheets.Add()
    zerosheet.Range(zerosheet.Cells(1,1),zerosheet.Cells(len(zero),len(zero[0]))).Value = zero
    zerobook.SaveAs(join(os.getcwd(),'zero.xls'))
    zerobook.Close()
    excel.Application.Quit()

def copyTimes(srcpath,x):
    global threads
    filelist = [f for f in listdir(txtLoc) if isfile(join(txtLoc, f))]
    numbs = len(filelist)
    div = round(numbs/x)
    count = 0
    leanmean(srcpath)

    pythoncom.CoInitialize()
    xl = win32.Dispatch('Excel.Application')
    xl_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, xl)
    for i in range(numbs):
            if i % div == 0:
                #threading goes here
                print("Creating Thread: ", count)
                shutil.copy('lean_sourcefile.xls','leanmean'+ str(count)+'.xls')
                count += 1
                starting = div*(count-1)
                ending = div*count -1 if div*count -1 < numbs else numbs-1
                t = threading.Thread(target = looper, args = (filelist[starting:ending],'leanmean'+ str(count-1)+'.xls'))
                threads.append(t)
                t.start()
                #print(starting, ending, numbs)
    for i in threads:
        i.join()


    writeZero()





copyTimes(getfilePath(),threadCount)


