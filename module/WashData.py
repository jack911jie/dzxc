import os
import sys
sys.path.append('i:/py/dzxc/module')
import days_calculate
import pandas as pd 
from datetime import datetime 
import copy

def crs_sig_table(xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\2020乐高课程签到表（周二）.xlsx'):
    df=pd.read_excel(xls,sheet_name='学生上课签到表',skiprows=1,header=None)
    df.iloc[0]=df.iloc[0].map(lambda x:str(x)[0:10])#将原有的时间格式变为字符串格式
    left_6=pd.Series(['机构','班级','姓名首拼','性别','ID','学生姓名','上课数量统计汇总'])
    # new_title=left_6.append(df.iloc[0].str.cat(df.iloc[1],sep=',')[6:]).tolist() #构建新的表头，使用了函数 df.iloc[0].str.cat
    title_time=left_6.append(df.iloc[0][7:]).tolist()
    title_crs=left_6.append(df.iloc[1][7:]).tolist()
    
    #学生实际上的课表
    df_std=df.iloc[2:]
    df_std_new=df_std.copy()
    df_std_new.columns=title_time

    for i in range(0,df_std_new.shape[0]):
        for j in range(6,df_std_new.shape[1]):
            if df_std_new.iloc[i,j]=='√':
                df_std_new.iloc[i,j]=title_crs[6:][j-6]
    
    #总课表
    df_crs_0=df.iloc[0:2,:].copy().T
    df_crs=df_crs_0.iloc[7:,:]
    df_crs.columns=['上课日期','课程名称']


    return {'total_crs':df_crs,'std_crs':df_std_new}

def std_term_crs(std_name='韦宇浠',start_date='20000927',end_date='21000105',xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\2020乐高课程签到表（周二）.xlsx'):
    df=crs_sig_table(xls=xls)    
    std_df=df['std_crs']
    std_name=std_name.strip()
    infos=std_df[std_df['学生姓名']==std_name]
    info_basic=infos[['机构','班级','姓名首拼','性别','ID','学生姓名','上课数量统计汇总']]
    info_crs_0=infos.iloc[:,7:]        
    info_crs=copy.copy(info_crs_0)
    info_crs.loc['aa']=info_crs_0.columns.values
    info_crs=info_crs.T
    info_crs.reset_index(drop=True,inplace=True)
    info_crs.columns=['课程名称','上课日期']
    # print(info_crs)
    info_crs['上课日期']=pd.to_datetime(info_crs['上课日期'])
    start_date=datetime.strptime(start_date[0:4]+'-'+start_date[4:6]+'-'+start_date[6:],'%Y-%m-%d')
    end_date=datetime.strptime(end_date[0:4]+'-'+end_date[4:6]+'-'+end_date[6:],'%Y-%m-%d')
    _std_crs=info_crs[(info_crs['上课日期']>=start_date) & (info_crs['上课日期']<=end_date)]
    std_crs=_std_crs.copy()
    std_crs['crs_name']=info_crs['课程名称']    
    std_crs.drop(labels='课程名称',axis=1,inplace=True)
    std_crs.columns=['上课日期','课程名称']
    # std_crs['上课日期'].apply(lambda x: x.strftime('%Y-%m-%d'))
    

    _total_crs=df['total_crs'].copy()
    _total_crs['上课日期']=pd.to_datetime(_total_crs['上课日期'])
    total_crs=_total_crs[(_total_crs['上课日期']>=start_date) & (_total_crs['上课日期']<=end_date)]
    # total_crs.dropna(inplace=True)


    return {'std_crs':std_crs,'total_crs':total_crs,'std_info':info_basic}

def std_feedback(std_name='韦宇浠',xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\每周课程反馈\\学员课堂学习情况反馈表.xlsx',weekday=2):
    wd='周'+days_calculate.num_to_ch(str(weekday))
    df=pd.read_excel(xls,sheet_name=wd,skiprows=1)
    df.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'姓名首拼','Unnamed: 2':'姓名','Unnamed: 3':'昵称','Unnamed: 4':'性别','Unnamed: 5':'优点特性','Unnamed: 6':'提升特性'},inplace=True)
    df_ability=df[['姓名','理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力','协作能力']]
    # print(df_ability)
    df_term_comment_txt=df.filter(regex='学期总结')
    df_term_comment=pd.concat([df[['ID','姓名']],df_term_comment_txt],axis=1)

    return {'df_ability':df_ability,'df_term_comment':df_term_comment}

def std_score(std_name='韦宇浠',start_date='20000927',end_date='21000105',xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\2020乐高课程签到表（周二）.xlsx'):
    #课堂积分计算
    def score(sht_name='课堂积分'):
        df_crs=pd.read_excel(xls,sheet_name=sht_name,skiprows=0,header=None)
        df_scores=df_crs.iloc[:,6:]
        crs_names=df_scores.iloc[0,:].str.replace(r'^\d{1,2}$','-',regex=True).dropna()
        std_scores=df_scores.iloc[2:,:]
        std_scores=std_scores.iloc[:,[x for x in range(3,std_scores.shape[1],4)]]
        pre_col_name=['未上课']*std_scores.shape[1]
        for k,v in enumerate(crs_names.values):
            pre_col_name[k]=v
        std_scores.columns=pre_col_name
        std_info=df_crs.iloc[2:,0:6]
        std_info.columns=df_crs.iloc[0,0:6].values    
        res=pd.concat([std_info,std_scores],axis=1)
        return res
   
    std_score_class=score(sht_name='课堂积分')
    std_score_activity=score(sht_name='活动积分')

    return {'std_score_class':std_score_class,'std_score_activity':std_score_activity}

def comments_after_class(crs_name_input,weekday,crs_list,crs_student,tch_cmt):
            crs_code=crs_name_input[0:4]
            crs_name=crs_name_input[4:]
            # print('正在读取学员和课程信息……',end='')
            df=pd.read_excel(crs_list) 
            crs=df.loc[df['课程编号']==crs_code]   
            knowledge=list(crs['知识点'])
            script=list(crs['课程描述'])
            dif_level=list(crs['难度'])
            instrument=list(crs['教具'])
            crs_info=[crs_name,knowledge[0],script[0],dif_level[0],instrument[0]]      
            stars=crs_info[-1].replace('*','★')
            crs_info[-1]=stars 
            
            df_stdInfo=pd.read_excel(crs_student,sheet_name='学生档案表')
            df_stdSig=pd.read_excel(crs_student,sheet_name='学生上课签到表',skiprows=2)
                     
            df_stdSig.rename(columns={'Unnamed: 0':'幼儿园','Unnamed: 1':'班级','Unnamed: 2':'姓名首拼', \
                                        'Unnamed: 3':'性别','Unnamed: 4':'ID','Unnamed: 5':'学生姓名', \
                                         'Unnamed: 6':'上年课时结余','Unnamed: 7':'购买课时', \
                                             'Unnamed: 8':'目前剩余课时','Unnamed: 9':'上课数量统计汇总'},inplace=True)
            # print(df_stdSig.columns)
            Students_sig=df_stdSig.loc[df_stdSig[crs_code+crs_name]=='√'][['幼儿园','班级','姓名首拼','学生姓名']] #上课的学生名单            
            Students=pd.merge(Students_sig,df_stdInfo,on='学生姓名',how='left') #根据学生名单获取学生信息
            Students_List=Students.values.tolist()

            NumtoC={'1':'一','2':'二','3':'三','4':'四','5':'五','6':'六','7':'日'}
            shtName='周'+NumtoC[str(weekday)]
            TeacherCmt=pd.read_excel(tch_cmt,sheet_name=shtName,skiprows=1)
            TeacherCmt.fillna('-',inplace=True)
            TeacherCmt.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'姓名首拼','Unnamed: 2':'学生姓名','Unnamed: 3':'昵称','Unnamed: 4':'性别','Unnamed: 5':'优点特性','Unnamed: 6':'缺点特性'},inplace=True)

            # print('完成')
            # print(Students_List)
            return {'std_list':Students_List,'crs_info':crs_info,'tch_cmt':TeacherCmt}    

# def term_summary_txt(std_name='韦宇浠',xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\每周课程反馈\\学员课堂学习情况反馈表.xlsx',weekday=2):
#     wd='周'+days_calculate.num_to_ch(str(weekday))
#     df=pd.read_excel(xls,sheet_name=wd,skiprows=1)


if __name__=='__main__':
    # print(std_feedback())
    # print(std_term_crs())
    # print(crs_sig_table())
    # print(std_score()['std_score_activity'])
    
    crs_list="/home/jack/data/大智小超/文档表格/课程信息表.xlsx"
    std_list="/home/jack/data/大智小超/文档表格/2020乐高课程签到表（周二）.xlsx"
    tch_cmt="/home/jack/data/大智小超/文档表格/每周课程反馈/学员课堂学习情况反馈表.xlsx"
    res=comments_after_class('L055螺旋桨飞机',weekday=2,crs_list=crs_list,crs_student=std_list,tch_cmt=tch_cmt)
    print(res['crs_info'])