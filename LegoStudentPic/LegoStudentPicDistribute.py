import os
import iptcinfo3
import logging
import json
import shutil
import pandas as pd
 
# logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.ERROR, format='%(levelname)s | %(funcName)s - 第 %(lineno)d 行 - %(message)s')
logger = logging.getLogger(__name__)

class LegoPics:
    def __init__(self,crsDate,crsName,weekday=2):
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'StudentsPicConfig.txt'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
                
            config=json.loads(_line)

        self.crsDate=str(crsDate)    
        self.crsName=crsName
        self.dir=config['乐高照片文件夹']
        self.stu_dir=config['乐高学员文件夹']
        self.stu_sigDir=config['2020乐高课程签到表文件夹']
        self.weekday=weekday

    def read_sig(self,weekday):
        if weekday==2:
            sigFile='2020乐高课程签到表（周二）.xlsx'
        elif weekday==6:
            sigFile='2020乐高课程签到表（周六）.xlsx'
        stdInfo=pd.read_excel(os.path.join(self.stu_sigDir,sigFile),sheet_name='学生上课签到表',skiprows=2)
        stdInfo.rename(columns={'Unnamed: 0':'幼儿园','Unnamed: 1':'班级','Unnamed: 2':'姓名首拼','Unnamed: 3':'性别','Unnamed: 4':'学生姓名','Unnamed: 5':'已上课数'},inplace=True)
        stdName=stdInfo['学生姓名'].tolist()
        stdPY=stdInfo['姓名首拼'].tolist()
        dictPY={}
        for i,v in enumerate(stdName):
            dictPY[v]=stdPY[i]

        return [dictPY,stdName]
        
    def dispatch(self):
        print('将打标签的照片分配到“I:\\每周乐高课_学员”中……')
        stdInfos=self.read_sig(weekday=self.weekday)
        dictPY=stdInfos[0]
        stdNamelist=stdInfos[1]
        for fn in os.listdir(os.path.join(self.dir,self.crsDate+'-'+self.crsName)):
            if fn[-3:].lower()=='jpg':
            #                 if iptcinfo3.IPTCInfo(os.path.join(self.dir,fn)):
                tag=code_to_str(iptcinfo3.IPTCInfo(os.path.join(self.dir,self.crsDate+'-'+self.crsName,fn)))      
                if len(tag)>0:
                    for _tag in tag:
                        if _tag in stdNamelist:
                            stu_dirName=os.path.join(self.stu_dir,self.crsDate+'-'+self.crsName,fn)
                            stu_pic_dirName=os.path.join(self.stu_dir,dictPY[_tag]+_tag)
                            if not os.path.exists(stu_pic_dirName):
                                os.makedirs(stu_pic_dirName)
                                oldName=os.path.join(self.dir,self.crsDate+'-'+self.crsName,fn)
                                newName=os.path.join(stu_pic_dirName,fn)
                                shutil.copyfile(oldName,newName)
                            else:
                                oldName=os.path.join(self.dir,self.crsDate+'-'+self.crsName,fn)
                                newName=os.path.join(stu_pic_dirName,fn)
                                shutil.copyfile(oldName,newName)
        

        print('\n完成')

            
            
def code_to_str(ss):
    s=ss['keywords']
    if isinstance(s,list):
        out=[]
        for i in s:
            out.append(i.decode('utf-8'))
    else:
        out=[ss.decode('utf-8')]
   
#     print(out)
    return out


if __name__=='__main__':
    stu_pics=LegoPics(crsDate=20200929,crsName='L033双翼飞机',weekday=6)
    stu_pics.dispatch()
