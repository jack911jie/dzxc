import os
import json
import pandas as pd
import iptcinfo3
from PIL import Image,ImageFont,ImageDraw
import logging
iptcinfo_logger=logging.getLogger('iptcinfo')
iptcinfo_logger.setLevel(logging.ERROR)


class pics:
    def __init__(self):
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'config.txt'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
            config=json.loads(_line)
        
        self.publicPicDir=config['公共图片']
        self.StdInfoDir=config['2019科学课签到表']
        self.CrsInfoDir=config['课程信息表']
        self.totalPics=config['照片总文件夹']
        # print(self.publicPicDir,self.CrsInfo)
        

    def putCover(self,stdName='韦成宇'):
        def read_excel():
            crsFile=['课程信息表.xlsx','课程信息']
            stdFile=['2019科学实验课学员档案2.xlsx','学员名单']
            crs=pd.read_excel(os.path.join(self.CrsInfoDir,crsFile[0]),skiprows=0,sheet_name=crsFile[1])
            stds=pd.read_excel(os.path.join(self.StdInfoDir,stdFile[0]),skiprows=2,sheet_name=stdFile[1])
            stds.rename(columns={'Unnamed: 0':'序号','Unnamed: 1':'姓名首拼','Unnamed: 2':'学生姓名','Unnamed: 3':'上课数量统计'},inplace=True)
            # std=stds[stds['学生姓名']==stdName]
            # std_basic=std[['姓名首拼','学生姓名']]
            # std_crs=std[std.iloc[:,:]=='√'].dropna(axis=1)
            # std_res=pd.concat([std_basic,std_crs],axis=1)
            # print(crs)
            return [crs,stds]

        def read_pics_old():
            lst=read_excel()
            crs,stds=lst[0],lst[1]
            stdList=stds['学生姓名'].tolist()
            # print(stdList)

            infos=[]
            for fileName in os.listdir(self.totalPics):
                fn=fileName.split('-')
                crsName=fn[3]
                real_addr=os.path.join(self.totalPics,fileName)
                tag=self.code_to_str(iptcinfo3.IPTCInfo(real_addr))
        
                if len(tag)>0:
                    for _tag in tag:
                        if _tag in stdList:
                            std_name=_tag
                knlg=crs[crs['实验主题']==crsName]['知识点'].tolist()
                infos.append([real_addr,std_name,crsName,knlg,fn[0]+'-'+fn[1]+'-'+fn[2]])
                
            return infos

        def putCover():
            infos=read_pics_old()
            infos=[infos[0]]
            for info in infos:
                img=Image.open(info[0])
                bg_h,bg_w=int(img.size[1]*0.2255),img.size[0]
                


        putCover()

    def code_to_str(self,ss):
        s=ss['keywords']
        if isinstance(s,list):
            out=[]
            for i in s:
                out.append(i.decode('utf-8'))
        else:
            out=[ss.decode('utf-8')]
    
        return out

if __name__=='__main__':
    pic=pics()
    pic.putCover()