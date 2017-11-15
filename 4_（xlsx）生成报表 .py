#!Python3
import os,openpyxl,shutil,time
os.chdir(os.getcwd())
date=time.strftime('%m-%d')
n1=date+'省级村村通总表.xlsx'
n2=date+'村村通自来水工程建设进度情况表.xlsx'
sourceDir=(os.path.abspath(os.path.join(os.getcwd(), "..",'2_省级村村通总表.xlsx')))
try:
    os.remove(n1)
except:
    pass
try:
    os.remove(n2)
except:
    pass
        
shutil.copy(sourceDir,n1)
print('复制目标文档……' )
print('开始读取文档……' )

wb1 = openpyxl.load_workbook(n1,data_only=True)
wb2 = openpyxl.load_workbook('样板.xlsx',data_only=True)

sheet1 = wb1.get_sheet_by_name('Sheet1')
sheet2 = wb2.get_sheet_by_name('Sheet1')


print('获取数据……' )
data=[]
data.append(['小计',20,'=SUM(C6:C25)',736241,14753,'=SUM(F6:F25)','=SUM(G6:G25)','=SUM(H6:H25)','=SUM(I6:I25)','=SUM(J6:J25)','=SUM(K6:K25)','=SUM(L6:L25)','=SUM(M6:M25)','=SUM(N6:N25)','=SUM(O6:O25)'])
list1=['X','Z','AB','AA','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL']
list2=['9','78','142','195','211','347','364','448','529','694','758','832','941','1012','1075','1182','1280','1343','1423','1439']
for n in list2:
    temp_list=[]
    temp_list.append(str(len(data)))
    for m in list1:
        cell=m+n
        temp_list.append(sheet1[cell].value)
    data.append(temp_list)

print('填写数据……' )

for r in range(5,26):
    for t in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O']:
        num=t+str(r)
        temp=data[r-5].pop(0)
        if temp!= 0:
            if t=='E' and r!=5:
                temp=round(temp)
            if t in ['G','I','K'] and r!=5:
                temp=round(temp,2)
            sheet2[num]=temp
 
sheet2['S5']='=O5+M5+K5+I5+G5'
sheet2['S6']='良垌加一减一'
sheet2['S7']='青平5个调整村庄共30.735元'    
print('处理结束……')

wb2.save(n2)
os.startfile(n2)
