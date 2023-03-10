import os
import sys
from numpy.lib.arraysetops import isin
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import WashData
import readConfig
import pandas as pd 
import re
import copy

class query:
    def __init__(self,place_input='001-超智幼儿园',wecom_id='1688856932305542'):
        config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'query.config'))
        self.std_in_class_list=config['学生分班表']
        self.std_in_class_list=self.std_in_class_list.replace('place',place_input).replace('$',wecom_id)
        self.std_info_dir=config['机构文件夹']
        self.std_info_dir=self.std_info_dir.replace('place',place_input).replace('$',wecom_id)


    def query_for_scores(self,std_input='',plus_tiyan='no'):
        df_score=WashData.std_all_scores(self.std_info_dir,plus_tiyan=plus_tiyan)
        # print(df_score)
        if isinstance(std_input,str):
            if re.match(r'w\d{3}',std_input):
                std_names_read=pd.read_excel(self.std_in_class_list)
                std_names=std_names_read[std_names_read['分班']==std_input]['学生姓名'].tolist()
                res=df_score[df_score['学生姓名'].isin(std_names)]
            else:
                if std_input=='' or std_input=='all':
                    res=df_score
                else:
                    std_names=[std_input]
                    res=df_score[df_score['学生姓名'].isin(std_names)]
        elif isinstance(std_input,list):
            std_names=std_input
            res=df_score[df_score['学生姓名'].isin(std_names)]
        else:
            res=''
            print('不能识别的输入类型')
        
        return res
    
    def query_for_std_score(self, std_name='DZ0034顾业熙',plus_tiyan='no'):
        df_std_crs_rec=pd.read_excel(os.path.join(self.std_info_dir,'学生档案',std_name+'.xlsx'),sheet_name='课程记录')
        df_std_score_verify=pd.read_excel(os.path.join(self.std_info_dir,'学生档案',std_name+'.xlsx'),sheet_name='积分兑换')

        if plus_tiyan=='yes':
            df_total_scr=df_std_crs_rec           
        else:
            df_total_scr=df_std_crs_rec[df_std_crs_rec['上课类型']=='正式']
        
        total_scr=df_total_scr['课堂积分'].sum()
        verify_scr=df_std_score_verify['兑换积分'].sum()
        remain_scr=total_scr-verify_scr

        df_res=pd.DataFrame({'学生姓名':std_name,'总积分':total_scr,'已兑换积分':verify_scr,'剩余积分':remain_scr},index=[0])

        return df_res

    def batch_query_for_score(self,term='2022秋',std_name='DZ0034顾业熙',plus_tiyan='no'):
        if isinstance(std_name,str):
            if std_name.startswith('w'):
                cls_dis=pd.read_excel(os.path.join(self.std_info_dir,'学生信息表','学生分班表.xlsx'),sheet_name='分班表')
                cls_dis['分班']=cls_dis['分班'].apply(lambda x: x.lower()) #确定为小写

                df_std_list=cls_dis[(cls_dis['学期']==term) & (cls_dis['分班']==std_name.lower())]
                df_std_list_new=copy.deepcopy(df_std_list)
                df_std_list_new['学生编码及姓名']=df_std_list['ID']+df_std_list['学生姓名']            
                std_list=df_std_list_new['学生编码及姓名'].tolist()
            else:
                std_list=[std_name]
        elif isinstance(std_name,list):
            std_list=std_name

        df_list=[]
        for std in std_list:
            df_res=self.query_for_std_score(std_name=std, plus_tiyan=plus_tiyan)
            df_list.append(df_res)
        
        df_all_std=pd.concat(df_list)
        df_all_std.reset_index(inplace=True)
        df_all_std.drop(['index'],axis=1,inplace=True)
        

        return df_all_std       


if __name__=='__main__':
    q=query()
    # res=q.query_for_s
    res=q.batch_query_for_score(std_name='DZ0032磨治丞',plus_tiyan='no')
    print(res)