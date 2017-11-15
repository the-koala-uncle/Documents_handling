#!Python
# -*- coding: utf-8 -*-

import os,sys,openpyxl
os.chdir(os.getcwd())
print('请稍等……' )
list_save=[]
list3_save=[]


##############################改动的地方######################################
name1='1.xlsx'
name2='2.xlsx'
wb1 = openpyxl.load_workbook(name1)
wb2 = openpyxl.load_workbook(name2)
sheet1 = wb1.get_sheet_by_name('Sheet1')
sheet2 = wb2.get_sheet_by_name('Sheet1')
#list3_save=['A','B','C','D','E','F','G','H','I','J','K','L','M','N''O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE']
#list3_save=['L','M']
list3_save=['A','B','C','D','E','F','G','H','I','J','K','L','M','N''O','P','Q','R','S','T','U']



##############################################################################

print('对比中……' )
start_row=1
end_row=29
print('******'.ljust(10)+name1[:-5].ljust(10)+name2[:-5])
for index in list3_save:
    for row in range(start_row,end_row+1):
        num =index + str(row)
        temp_value1= sheet1[num].value
        temp_value2= sheet2[num].value
        if temp_value1!=temp_value2:
            try:
                temp_value2=int(temp_value2)
                temp_value1=int(temp_value1)
                if not (temp_value1==None and temp_value2==0):
                    if not (temp_value2==None and temp_value1==0):
                        if abs(temp_value1-temp_value2)>1:
                            list_save.append(num)
                            try:
                                print(str(num).ljust(10)+str(temp_value1).ljust(10)+str(temp_value2))
                            except:
                                print('')
                                print(str(num))
                                print(temp_value1)
                                print(temp_value2)

            except:
                if not (temp_value1==None and temp_value2==0):
                    if not (temp_value2==None and temp_value1==0):
                            list_save.append(num)
                            try:
                                print(str(num).ljust(10)+str(temp_value1).ljust(10)+str(temp_value2))
                            except:
                                print('')
                                print(str(num))
                                print(temp_value1)
                                print(temp_value2)

#print(list_save)
print('*****************************************************')
print('不同之处共有……')
print(len(list_save))
print('处理结束……')
