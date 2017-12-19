#!Python
# -*- coding: utf-8 -*-

import os,sys,openpyxl
os.chdir(os.getcwd())
print('请稍等……' )

##############################改动的地方######################################
start_row=1
end_row=29
start_column='A'
end_column='AE'
error=0
#error=0.4
list3_save=['A','B','C','D','E','F','G','H','I','J','K','L','M','N''O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE']
##############################################################################

list_xlsx=[files for files in os.listdir('.') if files.endswith('.xlsx')]
if len(list_xlsx)!=2:
    sys.exit(1)
name1=list_xlsx[0]
name2=list_xlsx[1]
wb1 = openpyxl.load_workbook(name1)
wb2 = openpyxl.load_workbook(name2)
sheet1 = wb1.get_sheet_by_name(wb1.get_sheet_names()[0])
sheet2 = wb2.get_sheet_by_name(wb2.get_sheet_names()[0])

list_save=[]
list3_save=list3_save[list3_save.index(start_column):list3_save.index(end_column)]
print('对比中……' )
print('******'.ljust(10)+name1[:-5].ljust(30)+name2[:-5])
print('******'.ljust(10)+wb1.get_sheet_names()[0].ljust(30)+wb1.get_sheet_names()[0])
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
                        if abs(temp_value1-temp_value2)>error:
                            list_save.append(num)
                            try:
                                print(str(num).ljust(10)+str(temp_value1).ljust(30)+str(temp_value2))
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
                                print(str(num).ljust(10)+str(temp_value1).ljust(30)+str(temp_value2))
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
