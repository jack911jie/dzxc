import os
import sys
import copy
import pandas as pd
import numpy as np
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.max_rows', None)


class StudentData:
    def __init__(self,work_dir='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室',template_fn='学生课堂行为评分标准表.xlsx'):
        self.work_dir=work_dir
        self.standard_dir=os.path.join(self.work_dir,'001-超智幼儿园','SOP')
        self.standard_fn=os.path.join(self.standard_dir,template_fn)

    def read_template(self):
        df_template=pd.read_excel(self.standard_fn,sheet_name='行为及评分总表')
        return df_template

    # def std_mark(self,std_name='邓恩睿',tb_name='1',std_mark_fn='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\反馈表\\2022\\2022春-学生课堂行为记录表（周一）.xlsx'):
    def std_mark(self,std_name='邓恩睿',tb_name='1',std_mark_fn='C:\\Users\\jack\\Desktop\\副本2022春-学生课堂行为记录表（周一）.xlsx'):
        big_tbl=self.read_template()
        #分数=2的所有项
        mid_score_tbl=big_tbl[big_tbl['分数']==2]

        df_print_tbl=pd.read_excel(self.standard_fn,sheet_name='打印给老师内容')
        df_base_tbl=df_print_tbl[['环节','课堂项目','描述','打印编号','细分编码','分数']]
        # print(df_print_tbl)

        def one_table(tb_name=tb_name):
            df_std_one=pd.read_excel(std_mark_fn,sheet_name=tb_name,skiprows=3,usecols='D:S')
            df_std_one.rename(columns={'Unnamed: 3':'行为描述'},inplace=True)
            df_std=df_std_one[['行为描述',std_name]]

            df_res=pd.concat([df_base_tbl,df_std],axis=1)
            _df_std_out=df_res[['打印编号','行为描述','细分编码','分数',std_name]]
            df_std_out=copy.deepcopy(_df_std_out)
            df_std_out.fillna(0,inplace=True)

            return df_std_out

        def one_tbl_score():
            this_std_score=one_table(tb_name=tb_name)
            #筛选出标记有分数的项
            marked_std_score=this_std_score.drop(this_std_score.loc[this_std_score[std_name]==0].index)

            ins_tbl=mid_score_tbl.merge(marked_std_score,how='left',on='细分编码')
            ins_tbl['行为描述']=ins_tbl.apply(lambda x: x['行为描述_x'] if pd.isna(x['行为描述_y']) else x['行为描述_y'],axis=1)
            ins_tbl['分数']=ins_tbl.apply(lambda x: x['分数_y'] if x[std_name]==1 else x['分数_x'],axis=1)   
            ins_tbl['姓名']=std_name         

            #筛选出结果
            res_ins_tbl=ins_tbl[['姓名','环节','分类编码','分类名称','细分内容','行为描述','分数']]

            return res_ins_tbl

        return one_tbl_score()


if __name__=='__main__':
    p=StudentData(work_dir='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室',template_fn='学生课堂行为评分标准表.xlsx')
    res=p.std_mark()
    print(res)