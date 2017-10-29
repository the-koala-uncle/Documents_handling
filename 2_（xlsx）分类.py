#!Python3
import os,sys,openpyxl
os.chdir(os.getcwd())
print('请稍等……' )

##############################改动的地方######################################
name1='廉江市贫困村自然村（2017-6-23）.xlsx'
name2='广东省新时期精准扶贫相对贫困户数据汇总表（廉江2017-7-16）.xlsx'
name3='save.xlsx'
wb1 = openpyxl.load_workbook(name1)
wb2 = openpyxl.load_workbook(name2)

sheet1 = wb1.get_sheet_by_name('20170620160652')
sheet2 = wb2.get_sheet_by_name('贫困自然村情况')


list1_save=[]
list2_save=[]
list3_save=[]
list11_save=[]
list22_save=[]
list33_save=[]
top1=2
end1=657

top2=2
end2=659
a=0
b=0
c=0
##############################################################################

print('获取数据……' )

for row in range(top1,end1+1):
    num ='F'+ str(row)
    if int(sheet1[num].value)==1:
        list11_save.append(sheet1['E'+ str(row)].value)
        list1_save.append([sheet1['B'+ str(row)].value,sheet1['C'+ str(row)].value,sheet1['E'+ str(row)].value])
    elif int(sheet1[num].value)==2:
        list22_save.append(sheet1['E'+ str(row)].value)
        list2_save.append([sheet1['B'+ str(row)].value,sheet1['C'+ str(row)].value,sheet1['E'+ str(row)].value])
    else:
        list3_save.append([sheet1['B'+ str(row)].value,sheet1['C'+ str(row)].value,sheet1['E'+ str(row)].value])

        
print('-------------' )
print('红色数据……' )
print(list11_save[0])
print(list1_save[0])
print(list1_save[-1])
print('黑色红色数据……' )
print(list22_save[0])
print(list2_save[0])
print(list2_save[len(list2_save)-1])
print(list2_save[-1])
print('无法分类数据……' )
print(list3_save)
print('-------------' )

##############################################################################
print('开始对比分类……' )
for row in range(top2,end2+1):
    row=str(row)
    sheet2['L'+row]='='+'I'+row+'+'+'J'+row+'+'+'K'+row
    num2 ='E'+ row
    if (str(sheet2[num2].value) in list11_save) or(str(sheet2[num2].value) in list22_save):
        for num_temp in range(0,len(list11_save)):
            if str(list1_save[num_temp][2]) == str(sheet2['E'+ str(row)].value) and str(list1_save[num_temp][0]) == str(sheet2['B'+ str(row)].value) and str(list1_save[num_temp][1]) == str(sheet2['C'+ str(row)].value):
                data='I'+row
                sheet2[data]= 1
                a=a+1
        for num_temp in range(0,len(list22_save)):
            if str(list2_save[num_temp][2]) == str(sheet2['E'+ str(row)].value) and str(list2_save[num_temp][0]) == str(sheet2['B'+ str(row)].value) and str(list2_save[num_temp][1]) == str(sheet2['C'+ str(row)].value):
                data='J'+row
                sheet2[data]= 2
                b=b+1
    else:
        print('有无法被分类数据产生')
        print(row)
        data='K'+row
        sheet2[data]= 5
        c=c+1
##############################################################################
print('*****************************************************')
wb2.save('output.xlsx')
print(a)
print(b)
print(c)
#os.startfile('output.xlsx')
print('处理结束……')
