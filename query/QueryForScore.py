import os
import sys
from numpy.lib.arraysetops import isin
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import pandas as pd 
import re
import WashData

class query:
    def __init__(self):
        self.std_in_class_list='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\学生信息表\\学生分班表.xlsx'

    def query_for_scores(self,std_input=''):
        df_score=WashData.std_all_scores()
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
    q.query_for_scores(std_input=['陶盛挺','农雨蒙','黄建乐'])