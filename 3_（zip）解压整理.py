#!python3
import os,sys,zipfile,time
os.chdir (os.getcwd ())

###############################获取所有解压文件##########################################
list_zip=[files for files in os.listdir('.') if files.endswith('.zip')]
list_pdf=[files for files in os.listdir('.') if files.endswith('.pdf')]
list_doc=[files for files in os.listdir('.') if files.endswith('.doc')]
list_docx=[files for files in os.listdir('.') if files.endswith('.docx')]
#################################获得文件名序号##########################################

def get_name():
    temp_list=[]
    for path,dirs,files in os.walk('.'):
        for dirname in dirs:
            try:
                temp_num=int(dirname[:2])
                temp_list.append(temp_num)
            except:
                pass
    if max(temp_list)+1<10:
        num='0'+str(max(temp_list)+1)
    else:
        num=str(max(temp_list)+1)
    date=time.strftime('%m-%d')                   
    name=''.join([num,'_（',date,'收）'])
    return(name)


#############################解压文件####################################################
 
for zipname in list_zip:
    dir_name=get_name()+os.path.basename(zipname).rstrip('.zip').rstrip('.rar')
    print(dir_name)
    
    z = zipfile.ZipFile(zipname,"r")
    z.extractall('name_invaluable_in_chinese')
    z.close()

    def renamedir(dirpath):
        d=dirpath
        dd=os.path.join(os.path.dirname(d),os.path.basename(d).encode('cp437').decode('gbk'))
        os.rename(d,dd)
        for i in os.listdir(dd):
            ipath=os.path.join(dd,i)
            if os.path.isdir(ipath):
                renamedir(ipath)
            else:
                os.rename(os.path.join(dd,i),os.path.join(dd,i.encode('cp437').decode('gbk')))
    renamedir('name_invaluable_in_chinese')
    os.rename('name_invaluable_in_chinese',dir_name)
    os.unlink(zipname)

#############################整理完成####################################################
for docname in list_doc:
    doc_name=get_name()+os.path.basename(docname)
    os.makedirs(doc_name)
    shutil.move(docname,doc_name)   

for docxname in list_docx:
    docx_name=get_name()+os.path.basename(docxname)
    os.makedirs(docx_name)
    shutil.move(docxname,docx_name)   
for pdfname in list_pdf:
    pdf_name=get_name()+os.path.basename(pdfname)
    os.makedirs(pdf_name)
    shutil.move(pdfname,pdf_name)   

_=len(list_zip)+len(list_doc)+len(list_docx)+len(list_pdf)
if _==0:
    print('不存在zip文件')
    input()
else:
    print('整理完成,共处理了',_,'个文件。',sep='')
    time.sleep(1)
