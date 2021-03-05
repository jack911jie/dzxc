import os
import sys
sys.path.append('i:/py/dzxc/module')
import days_calculate
import pandas as pd 

def crs_sig_table(xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\2020乐高课程签到表（周二）.xlsx'):
    df=pd.read_excel(xls,sheet_name='学生上课签到表',skiprows=1,header=None)
    df.iloc[0]=df.iloc[0].map(lambda x:str(x)[0:10])#将原有的时间格式变为字符串格式
    left_6=pd.Series(['机构','班级','姓名首拼','性别','学生姓名','上课数量统计汇总'])
    # new_title=left_6.append(df.iloc[0].str.cat(df.iloc[1],sep=',')[6:]).tolist() #构建新的表头，使用了函数 df.iloc[0].str.cat
    title_time=left_6.append(df.iloc[0][6:]).tolist()
    title_crs=left_6.append(df.iloc[1][6:]).tolist()
    df_std=df.iloc[2:]
    df_std_new=df_std.copy()
    df_std_new.columns=title_time

    for i in range(0,df_std_new.shape[0]):
        for j in range(6,df_std_new.shape[1]):
            if df_std_new.iloc[i,j]=='√':
                df_std_new.iloc[i,j]=title_crs[6:][j-6]
    
    # df_std_new.to_excel('e:/temp/kkkkkk.xlsx')
    # print(df_std_new)
    return df_std_new

def std_feedback_ability(xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\每周课程反馈\\学员课堂学习情况反馈表.xlsx',weekday=2):
    wd='周'+days_calculate.num_to_ch(str(weekday))
    df=pd.read_excel(xls,sheet_name=wd,skiprows=1)
    df.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'姓名首拼','Unnamed: 2':'姓名','Unnamed: 3':'昵称','Unnamed: 4':'性别','Unnamed: 5':'优点特性','Unnamed: 6':'提升特性'},inplace=True)
    df_ability=df[['姓名','理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力','协作能力']]
    # print(df_ability)
    return df_ability



if __name__=='__main__':
    std_feedback_ability()