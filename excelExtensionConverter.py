import win32com.client as win32
import os
from win32com.client import Dispatch


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)


def convertExcel(targetExt, getFile, saveFile):
    if targetExt == '.xls':
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(getFile)
        wb.SaveAs(saveFile + "x", FileFormat=51)
        wb.Close()
        excel.Application.Quit()
    elif targetExt == '.xlsx':
        excel = Dispatch('Excel.Application')
        wb = excel.Workbooks.Add(getFile)
        wb.SaveAs(saveFile[:-1], FileFormat=56)
        excel.Quit()
    else:
        pass


convertPath = input("경로 입력 :")
savePath = convertPath + '_change'
createFolder(savePath)

fileList = os.listdir(convertPath)
for fileName in fileList:
    file, targetExt = os.path.splitext(fileName)
    getFile = convertPath + '\\' + fileName
    saveFile = savePath + '\\' + fileName
    print('Success Save File : ' + saveFile)
    convertExcel(targetExt, getFile, saveFile)
