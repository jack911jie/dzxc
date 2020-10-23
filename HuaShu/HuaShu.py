import os
import pandas as pd
import json

class huashu:
    def __init__(self):
        with open (os.path.join(os.getcwd(),'HuaShu','HuaShu.config'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            
            _line=''
            for line in lines:
                _line=_line+line
            
            config=json.loads(_line)
            self.StdSig=config['学员签到表']
            self.CrsTable=config['课程信息表']
            self.expDir=config['话术导出文件夹']
            self.talkModule='亲爱的家长：\n这是{0}本周科学机器人课《{1}》的课堂反馈，本节课的主要学习内容是：\n{2}\n'
            
        
    def legoTalk(self,crsname):
        stds=pd.read_excel(self.StdSig,skiprows=2)
        stds.dropna(axis=1,how='all',inplace=True)
        stds.rename(columns={'Unnamed: 0':'幼儿园','Unnamed: 1':'班级','Unnamed: 2':'姓名首拼','Unnamed: 3':'性别','Unnamed: 4':'姓名'},inplace=True)
#         stdIn=stds[stds[crsname]=='√'][['姓名首拼','姓名',crsname]]
        stdIn=stds[['姓名首拼','姓名',crsname]]
        
        crs=pd.read_excel(self.CrsTable)
        crs_klg=crs[crs['课程名称']==crsname]['知识点'].tolist()[0]

        stdIn=stdIn.assign(话术='-')
        
        for std in stdIn['姓名'].tolist():
            txt=self.talkModule.format(std,crsname,crs_klg)
            k=stdIn['话术'][(stdIn['姓名']==std) & (stdIn[crsname]=='√')]=txt
        out=stdIn[['姓名首拼','姓名','话术']]
#         print(stdIn)
        
        writer = pd.ExcelWriter(os.path.join(self.expDir,'导出话术.xlsx'))
        out.to_excel(writer)
        writer.save()
        print('完成')
        
if __name__=='__main__':
    talk=huashu()
    talk.legoTalk('啃骨头的小狗')
    