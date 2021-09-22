import os
import sys
from numpy.lib.arraysetops import isin
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import WashData
import readConfig
import days_calculate
import pandas as pd 
import re

class Query:
    def __init__(self,place_input='001-超智幼儿园'):
        config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'query.config'))
        self.std_in_class_list=config['学生分班表']
        self.std_in_class_list=self.std_in_class_list.replace('$',place_input)
        self.std_info_dir=config['机构文件夹']
        self.std_info_dir=self.std_info_dir.replace('$',place_input)

    def std_class_taken(self,weekday=[6],display='print_list',format='only_clsname'):
        wds=[]
        for wkdy in weekday:
            wds.append(days_calculate.num_to_ch(wkdy))
        cls=[]
        fn_dir=os.path.join(self.std_info_dir,'学生信息表')
        for xls in os.listdir(fn_dir):
            if re.match(r'\d\d\d\d\w{1}-.*.xlsx$',xls) and xls[-8:-6]!='体验':
                for wd in wds:
                    if xls[-7:-6]==wd:
                        fn=os.path.join(fn_dir,xls)
                        info_taken=WashData.class_taken(fn)
                        cls.extend(info_taken['cls_taken'])
        cls_only_name=[item[9:] for item in cls]


        if format=='only_clsname':
            cls=cls_only_name
        
        if display=='print_list':
            for line in cls:
                print(line)
        
        return cls

    def check_repeat(self,term='2021秋',weekday=5,crs_name='L103摩托车骑士'):
        wd=days_calculate.num_to_ch(weekday)
        fn_dir=os.path.join(self.std_info_dir,'学生信息表')
        fn=os.path.join(fn_dir,term+'-学生信息表（周'+wd+'）.xlsx')
        df_current=pd.read_excel(fn,skiprows=1)
        std_list=[item[0] for item in df_current.iloc[:,4:5].values.tolist()]
        info_current=[crs_name,std_list]
        rpt=0
        rpt_info=[]
        for xls in os.listdir(fn_dir):
            if re.match(r'\d\d\d\d\w{1}-.*.xlsx$',xls) and xls[-8:-6]!='体验':
                fn_ever=os.path.join(fn_dir,xls)
                info_taken=WashData.class_taken(fn_ever)
                cls_taken_get=[cls[9:] for cls in info_taken['cls_taken']]
                #输入的课程名单是否在既往的课程中
                
                if crs_name in cls_taken_get:
                    #当前学期的学生姓名是否在过去的名单中
                    for std in std_list:
                        if std in info_taken['std_names']:
                            rpt_info.append([std,crs_name,xls[:5],xls[-8:-6]])
                            rpt+=1
                
        if rpt>0:
            print('\n在 {} 学期 星期{} 的学生课程中检测到以下重复\n'.format(term,wd))
            for info in rpt_info:
                print(info)
        else:
            print('未检测到重复值')



if __name__=='__main__':
    qry=Query(place_input='001-超智幼儿园')
    # qry.std_class_taken(weekday=[1,4],display='print_list',format='only_clsnam')
    qry.check_repeat(term='2021秋',weekday=1,crs_name='L042玉免捣药') 