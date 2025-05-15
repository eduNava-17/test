import xlwings as xw
import pandas as pd
import numpy as np
import math

file1 = r'D:\Substitucion_1.xlsx'
ds1 = xw.Book(file1).sheets['Sheet1'] #
ds2 = xw.Book(file1).sheets['Sheet2'] # put here file2 and the sheet you want to compare with file1(Sheet1)
abc = ['A','B']
data1 = []
data2 = []
start_header = 2
for i in range(len(abc)):
    data1.append([x for x in ds1.range('{}{}:{}10'.format(abc[i],start_header,abc[i])).value])
    data2.append([x for x in ds2.range('{}{}:{}10'.format(abc[i],start_header,abc[i])).value])

nones = []
for num, val in enumerate(data1[0]):
    if val == None:
        nones.append([num + start_header,data1[1][num]])

set2 = set(data2[1])
set1 = np.array(nones).T[1]

elementos = []
for num, val in enumerate(data2[1]):
    for val2 in set1:
        if val in set1:
            if val != None:
                elementos.append([num + start_header,data2[0][num]])
                break

elementos = np.array(elementos)
# print (elementos)
# common_elements = list(set1.intersection(set2)) #useful

for i in range(len(elementos)):
    ds1.cells(int(elementos.T[0][i]),1).value = elementos.T[1][i]
    ds1.cells(int(elementos.T[0][i]),1).color = (0,100,255)
    



