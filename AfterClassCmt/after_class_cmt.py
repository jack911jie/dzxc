from email.mime import application
import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import days_cal
from datetime import datetime
import random

import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.max_rows', None)




class AfterClass:
    def __init__(self,place='001-超智幼儿园',work_dir='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室'):
        self.work_dir=work_dir
        self.place_dir=os.path.join(work_dir,place)

    def std_score(self,std_name='邓恩睿',weekday=1,term='2022春',crs_name='20220411-L148小小直升机'):    
        self.feedback_dir=os.path.join(self.place_dir,'每周课程反馈','反馈表',term[:4])
        wd=days_cal.num_to_weekday(weekday)
        df_crs_info=pd.read_excel(os.path.join(self.feedback_dir,term+'-学生课堂行为记录表（周'+wd+'）.xlsx'),sheet_name='课程信息表',skiprows=1)
        crs_num=df_crs_info[df_crs_info['课程名称']==crs_name]['课时'].tolist()[0]
        df_score=pd.read_excel(os.path.join(self.feedback_dir,term+'-学生课堂行为记录表（周'+wd+'）.xlsx'),sheet_name=str(crs_num),skiprows=3)
        df_score.rename(columns={'Unnamed: 0':'环节','Unnamed: 1':'能力','Unnamed: 2':'序号','Unnamed: 3':'表现行为'},inplace=True)
        df_std_score=df_score[['环节','能力','序号','表现行为',std_name]]
        return df_std_score
    
    def std_complete_score(self,std_name='邓恩睿',weekday=1,term='2022春',crs_name='20220411-L148小小直升机',
                            ref_table='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\SOP\\学生课堂行为评分标准表00.xlsx'):
        #读取学生分数,筛选出打有分数的项目
        df_std_score=self.std_score(std_name=std_name,weekday=weekday,term=term,crs_name=crs_name)
        df_std_score=df_std_score[~pd.isna(df_std_score[std_name])]

        #读取评语库
        cmt_lib_num=ref_table[-7:-5]
        cmt_lib_fn=os.path.join(self.place_dir,'SOP','学生评语库.xlsx')
        df_cmt_all=pd.read_excel(cmt_lib_fn,sheet_name='评语库'+cmt_lib_num)
        df_cmt_mid=df_cmt_all[df_cmt_all['分数']==2]
        
        #读取打印版本
        df_print=pd.read_excel(os.path.join(self.place_dir,'SOP',ref_table),sheet_name='打印给老师内容')

        #生成打有分数的项目的打印要亮码及细分编码
        df_cross=df_print.merge(df_std_score,how='inner',on='序号')
        df_mark=df_cross[['序号','打印编码','细分编码','分数','描述']]

    

        df_pre_merge=df_cmt_mid.merge(df_mark,how='outer',on='细分编码')

        df_pre_merge['打印编码']=df_pre_merge.apply(lambda x: x['打印编码_x'] if pd.isna(x['打印编码_y']) else x['打印编码_y'],axis=1)
        df_pre_merge['行为描述']=df_pre_merge.apply(lambda x: x['行为描述'] if pd.isna(x['描述']) else x['描述'],axis=1)
        df_pre_merge['分数']=df_pre_merge.apply(lambda x: x['分数_x'] if pd.isna(x['分数_y']) else x['分数_y'],axis=1)
        df_pre_merge['学生评语']=df_pre_merge['打印编码'].apply(lambda x: self.sel_cmt(df_cmt_all=df_cmt_all,prt_code=x))

        df_cmt_res=df_pre_merge[['打印编码','细分编码','行为描述','分数','学生评语']]
        
        return df_cmt_res

    

    def sel_cmt(self,df_cmt_all,prt_code='01KQ12'):
        cmt=df_cmt_all[df_cmt_all['打印编码']==prt_code]['评语'].tolist()        
        out_cmt=random.choice(cmt)
        # print(out_cmt)
        return out_cmt



if __name__=='__main__':
    p=AfterClass(place='001-超智幼儿园',work_dir='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室')
    p.std_complete_score(std_name='邓恩睿',weekday=1,term='2022春',crs_name='20220411-L148小小直升机',
                            ref_table='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\SOP\\学生课堂行为评分标准表00.xlsx')
