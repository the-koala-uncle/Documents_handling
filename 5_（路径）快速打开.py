import os,sys
from collections import OrderedDict
print('请输入路径代号：')


filepath= OrderedDict()
filepath['spb']=['审批表','D:/村村通/省级村村通/备查及验收/村村通申请资金/省级村村通项目申请资金']
filepath['zjbf']=['资金拨付表(财局）','D:\村村通\省级村村通\备查及验收\村村通申请资金']
filepath['gzsb']=['各镇上报','D:/村村通/省级村村通/上报内容/月报/各镇上报/2017年']
    
filepath['zjyb']=['湛江月报','D:/村村通/省级村村通/上报内容/月报/湛江市表/新表报表']
filepath['zjys']=['湛江验收批次','D:/村村通/省级村村通/备查及验收/湛江验收']

filepath['tbb']=['通报表','D:/村村通/省级村村通/上报内容/月报/廉江市表/廉江市村村通自来水工程任务完成情况汇总表.xls']
filepath['tbw']=['通报文','D:\村村通\省级村村通\文件通知（发文）/2_村村通信息通报']
filepath['tbqm']=['通报签名','D:\局文件、报帐表等\公文处理\廉江市水务局网站信息发布审批表.doc']

    
filepath['syts']=['所有通水','D:\村村通\省级村村通\上报内容\月报\湛江市表']
filepath['fp']=['精准扶贫','D:\村村通\省级村村通\上报内容\其他\其他报表\精准扶贫']
    
filepath['zdx']=['重点县','D:\重点县及水系连通\重点县\重点县2014年/2014年度验收资料']
filepath['esh']=['20户以上自然村采集表','D:\村村通\省级村村通\上报内容\表格/2017年/20户以上村庄新增村庄各镇确认表(青平、河唇、石颈、城北，安铺没报）']

filepath['直接回车']=['邮箱','']

for _ in filepath:
    print(filepath[_][0].ljust(15,'-')+ _)
path=str(input())
if path in filepath.keys():
    os.startfile(filepath[path][1].replace('\\','/') )
if path==''or path=='yx':
    os.system('"C:/Users\Administrator/AppData/Local/360Chrome/Chrome/Application/360chrome.exe" mail.163.com')
    sys.exit(1)


