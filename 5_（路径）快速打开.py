import os,sys
from collections import OrderedDict
print('请输入路径代号：')


filepath= OrderedDict()
filepath['bb']=['报表生成','D:\文件储存\~_文件处理区']

filepath['spb']=['审批表','D:/村村通/省级村村通/备查及验收/村村通申请资金/省级村村通项目申请资金']
filepath['zjbf']=['资金拨付表(财局）','D:\村村通\省级村村通\备查及验收\村村通申请资金']
filepath['gzsb']=['各镇上报','D:/村村通/省级村村通/上报内容/月报/各镇上报/2017年']
    
filepath['zjyb']=['湛江月报','D:/村村通/省级村村通/上报内容/月报/湛江市表/新表报表']
filepath['zjys']=['湛江验收批次','D:/村村通/省级村村通/备查及验收/湛江验收']

filepath['tbb']=['通报表','D:/村村通/省级村村通/上报内容/月报/廉江市表/廉江市村村通自来水工程任务完成情况汇总表.xls']
filepath['tbw']=['通报文','D:\村村通\省级村村通\文件通知（发文）/2_村村通信息通报']
filepath['tbqm']=['通报签名','D:\局文件、报帐表等\公文处理\廉江市水务局网站信息发布审批表.doc']

    
filepath['syts']=['所有通水','D:\村村通\省级村村通\上报内容\月报\湛江市表/湛江市村村通自来水工程项目县（市、区）通自来水现状调查统计表(人数）.xlsx']
filepath['fp']=['精准扶贫','D:\村村通\省级村村通\上报内容\其他\其他报表\精准扶贫']

filepath['rw']=['任务到村明细表','D:\村村通\省级村村通\备查及验收\村村通备查\任务、资金/任务到村明细（修改2016.11.18）.xls']
filepath['zdx']=['重点县','D:\重点县及水系连通\重点县\重点县2014年/2014年度验收资料']
filepath['nyrw']=['廉江市农村饮水任务明细表','D:\农村饮水\新建文件夹\台帐/廉江市农村饮水安全人口一览表.xls']
filepath['nyhz']=['农村饮水现状普查表','D:\村村通\省级村村通\上报内容\月报\湛江市表\台帐/复件 汇总.xls']
filepath['esh']=['自然村采集表(20户以上)','D:\村村通\省级村村通\上报内容\表格/2017年/20户以上村庄新增村庄各镇确认表(青平、河唇、石颈、城北，安铺没报）']



for _ in filepath:
    print(filepath[_][0].ljust(15,'-')+ _)
print()
print('村村通总表------（直接回车）')
path=str(input())
if path in filepath.keys():
    os.startfile(filepath[path][1].replace('\\','/') )
if path=='':
    os.startfile('D:/文件储存/2_省级村村通总表.xlsx')
##if path==''or path=='yx':
##    os.system('"C:/Users\Administrator/AppData/Local/360Chrome/Chrome/Application/360chrome.exe" mail.163.com')
##    sys.exit(1)
