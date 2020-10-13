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
    def __init__(self,crsName):
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'config.txt'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
                
            config=json.loads(_line)
            
        self.crsName=crsName
        self.dir=config['乐高照片文件夹']
        self.stu_dir=config['乐高学员文件夹']
        self.stu_sig=config['2020乐高课程签到表']

    def read_sig(self):
        stdInfo=pd.read_excel(self.stu_sig,sheet_name='学生上课签到表',skiprows=2)
        stdInfo.rename(columns={'Unnamed: 0':'幼儿园','Unnamed: 1':'班级','Unnamed: 2':'姓名首拼','Unnamed: 3':'性别','Unnamed: 4':'学生姓名','Unnamed: 5':'已上课数'},inplace=True)
        stdName=stdInfo['学生姓名'].tolist()
        stdPY=stdInfo['姓名首拼'].tolist()
        dictPY={}
        for i,v in enumerate(stdName):
            dictPY[v]=stdPY[i]

        return [dictPY,stdName]
        
    def dispatch(self):
        print('将打标签的照片分配到“I:\\每周乐高课_学员”中……')
        for fn in os.listdir(os.path.join(self.dir,self.crsName)):
            if fn[-3:].lower()=='jpg':
#                 if iptcinfo3.IPTCInfo(os.path.join(self.dir,fn)):
                stdInfos=self.read_sig()
                dictPY=stdInfos[0]
                stdNamelist=stdInfos[1]
                tag=code_to_str(iptcinfo3.IPTCInfo(os.path.join(self.dir,self.crsName,fn)))                
                if len(tag)>0:
                    for _tag in tag:
                        if _tag in stdNamelist:
                            stu_dirName=os.path.join(self.stu_dir,self.crsName,fn)
                            stu_pic_dirName=os.path.join(self.stu_dir,dictPY[_tag]+_tag)
                            if not os.path.exists(stu_pic_dirName):
                                os.makedirs(stu_pic_dirName)
                                oldName=os.path.join(self.dir,self.crsName,fn)
                                newName=os.path.join(self.stu_dir,_tag,fn)
                                shutil.copyfile(oldName,newName)
                            else:
                                oldName=os.path.join(self.dir,self.crsName,fn)
                                newName=os.path.join(self.stu_dir,_tag,fn)
                                shutil.copyfile(oldName,newName)
        

        print('\n……完成')

            
            
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
    stu_pics=LegoPics('20201013夏天的手摇风扇')
    stu_pics.dispatch()
