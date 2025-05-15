import xlwings as xw
import pandas as pd
import numpy as np
import math

file1 = r'D:\Substitucion_1.xlsx'
ds1 = xw.Book(file1).sheets['Sheet1'] #
ds2 = xw.Book(file1).sheets['Sheet2'] # put here file2 and the sheet you want to compare with file1(Sheet1)
ds1_pd = pd.read_excel(file1, sheet_name = 'Sheet1')
ds2_pd = pd.read_excel(file1, sheet_name = 'Sheet2')

my_list11 = ds1_pd['Header1'].tolist()
my_list12 = ds1_pd['Header2'].tolist()
my_list21 = ds2_pd['Header1'].tolist()
my_list22 = ds2_pd['Header2'].tolist()

start_header = 2

nones = []
for num, val in enumerate(my_list11):
    if pd.isnull(val):
        nones.append([num + start_header,my_list12[num]])

set2 = set(my_list22)
set1 = np.array(nones).T[1]

elementos = []
for num, val in enumerate(my_list22):
    for val2 in set1:
        if val in set1:
            if val != None:
                elementos.append([num + start_header,my_list21[num]])
                break

elementos = np.array(elementos)
# print (elementos)

# common_elements = list(set1.intersection(set2)) #useful

for i in range(len(elementos)):
    ds1.cells(int(elementos.T[0][i]),1).value = elementos.T[1][i]
    ds1.cells(int(elementos.T[0][i]),1).color = (255,255,0)
    




