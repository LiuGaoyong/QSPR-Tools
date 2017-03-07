#!/usr/bin/py3
# -*- coding:utf-8 -*-

#Author: Young Liu

import pprint as pp
import os
import xlrd
import xlwt
import subprocess as sp
from datetime import date,datetime

# get the result of a shell command/script
def getShellResult(com='', obj=''):
    output = sp.Popen([com, obj], stdout=sp.PIPE, shell=True)
    return output.communicate()[0][:-1]

#test
#print(getShellResult('/home/jun/lgy-WorkSpace/molDescrip/molMass /home/jun/lgy-WorkSpace/molDescrip/10.log'))

# get the all files list
def folder2List(dirPath):
    fileList=[]
    for root,dirs,files in os.walk(dirPath):    #root is always fileObj's father path
        for fileObj in files:                   #root 永远是 fileObj 的父路径
            fileList.append(os.path.join(root,fileObj))
    return fileList

#test
#a=folder2List(input('input the path of your Gaussian outfiles: '))
#pp.pprint(len(a), a)

def getDataMatrix():
    comPath = r'/home/jun/lgy-WorkSpace/molDescrip/'
    objFileList = folder2List(input('input the path of your Gaussian outfiles: '))
    matrix = []
    for i in objFileList:
        #row0 = [u'FileName', u'MolFormula', u'MolMass', u'E_total']
        FN = i.split('/')[-1]
        MF = ''
        MM = float(getShellResult(comPath+'molMass  '+i))
        ET = float(getShellResult(comPath+'totEnergy  '+i))
        matrix.append([FN, MF, MM, ET])
    return matrix

#test
#pp.pprint(getDataMatrix())

def set_style(name, height, bold=False): #for main()
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

def main():
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row0 = [u'FileName', u'MolFormula', u'MolMass', u'E_total']
    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i], set_style('Time New Roman', 220, True))
    
    dataMatrix = getDataMatrix()
    for j in range(0,len(dataMatrix)):
        for i in range(0, len(dataMatrix[j])):
            sheet1.write(j+1, i, dataMatrix[j][i], set_style('Time New Raman', 220))
    f.save('demo.xls')

if __name__ == "__main__" :
    main()
