#!Python3
import os,openpyxl,shutil,time
os.chdir(os.getcwd())

print('请选择命令：')
print('1.生成村村通进度表')
print('2.生成通报表')
index =int(input())
date=time.strftime('%m-%d@%H')
n1=date+'省级村村通总表.xlsx'
n2=date+'村村通自来水工程建设进度情况表.xlsx'
n3=date+'通报表.xlsx'
sourceDir=(os.path.abspath(os.path.join(os.getcwd(), "..",'2_省级村村通总表.xlsx')))
try:
    os.remove(n1)
except:
    pass
try:
    os.remove(n2)
except:
    pass
try:
    os.remove(n3)
except:
    pass    
shutil.copy(sourceDir,n1)
print('复制目标文档……' )
print('开始读取文档……' )

if index==1:

    wb1 = openpyxl.load_workbook(n1,data_only=True)
    wb2 = openpyxl.load_workbook('模板1.xlsx',data_only=True)

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

########################################################
if index==2:
    wb1 = openpyxl.load_workbook(n1,data_only=True)
    wb3 = openpyxl.load_workbook('模板2.xlsx',data_only=True)

    sheet1 = wb1.get_sheet_by_name('Sheet1')
    sheet3 = wb3.get_sheet_by_name('Sheet1')


    print('获取数据……' )
    data2=[]
    list2=['9','78','142','195','211','347','364','448','529','694','758','832','941','1012','1075','1182','1280','1343','1423','1439']
    for n in list2:
        temp_list=[]
        n=str(n)
        temp_list.append(sheet1['X'+n].value)
        temp_list.append(sheet1['Z'+n].value)
        finish=sheet1['AG'+n].value+sheet1['AI'+n].value+sheet1['AK'+n].value
        built=finish+sheet1['AE'+n].value
        temp_list.append(built)
        temp_list.append(finish)
        temp_list.append(sheet1['AB'+n].value)
        peo1=sheet1['L'+n].value
        peo2=0
        for m_n in range(0,19,2):
            num_1=int(list2[m_n])+1
            num_2=int(list2[m_n+1])
            for nn in range(num_1,num_2):
                try:
                    int(sheet1['AG'+nn].value)
                    try:
                        int(sheet1['AM'+nn].value)
                    except:
                        peo2=peo2+int(sheet1['G'+nn].value)
                except:
                    pass
        peo=int(peo1)+peo2

        temp_list.append(peo)
        fate=sheet1['AH'+n].value+sheet1['AJ'+n].value+sheet1['AL'+n].value
        temp_list.append(fate)    
        temp_list.append(round(peo/sheet1['AB'+n].value,4))
        data2.append(temp_list)
    data2.sort(key=lambda x:x[-1],reverse=True)


    print('填写数据……' )

    for r in range(5,25): 
        for t in ['B','C','D','E','G','H','K']:
            num=t+str(r)
            temp=data2[r-5].pop(0)
            sheet3[num]=temp

    wb3.save(n3)
    os.startfile(n3)

