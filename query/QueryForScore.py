import os
import sys
from numpy.lib.arraysetops import isin
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import WashData
import readConfig
import pandas as pd 
import re

class query:
    def __init__(self,place_input='001-超智幼儿园'):
        config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'query.config'))
        self.std_in_class_list=config['学生分班表']
        self.std_in_class_list=self.std_in_class_list.replace('$',place_input)
        self.std_info_dir=config['机构文件夹']
        self.std_info_dir=self.std_info_dir.replace('$',place_input)

    def query_for_scores(self,std_input=''):
        df_score=WashData.std_all_scores(self.std_info_dir)
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

        


if __name__=='__main__':
    q=query()
    res=q.query_for_scores(std_input='w101')
    print(res)