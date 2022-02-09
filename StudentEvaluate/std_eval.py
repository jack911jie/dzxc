import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),'modules'))
from readConfig import readConfig
import pandas as pd 

class StdEval:
    def __init__(self,place='001-超智幼儿园'):
        config=readConfig(os.path.join(os.path.dirname(__file__),'configs','StudentEvalute.config'))
        self.std_info_dir=config['学生信息表文件夹'].replace('#',place)
        self.std_comment_dir=config['课程反馈文件夹'].replace('#',place)
        self.sop_dir=config['SOP文件夹'].replace('#',place)
        self.std_eval_dir=config['学生课堂评分表文件夹'].replace('#',place)

    def read_data(self,fn='D:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\学生课堂评分表\\2022春\\20220209_L140野马战斗机_学生课堂行为评分表.xlsx',
                    fn_crs_info='D:\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\2-乐高课程\\课程信息表.xlsx'):
        df_crs=pd.read_excel(fn_crs_info,sheet_name='课程信息')
        df_cmt=pd.read_excel(fn,skiprows=3)
        df_std_cmt=df_cmt.iloc[:,[1,2,6,7,8,9,]].dropna(axis=1,how='all')
        df_std_cmt.fillna(2,inplace=True)
        print(df_std_cmt)


if __name__=='__main__':
    p=StdEval()
    p.read_data()
