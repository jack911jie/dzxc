import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),'modules'))
import days_calculate
import pandas as pd 
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
import numpy as np
from datetime import datetime 
import copy
import re

# def crs_sig_table(xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生信息表\\2020秋-学生信息表（周二）.xlsx'):
def crs_sig_table(xls='E:\\temp\\2021春-学生信息表（周四）_test.xlsx'):
    df=pd.read_excel(xls,sheet_name='学生上课签到表',skiprows=1,header=None)
    # df.iloc[0]=df.iloc[0].map(lambda x:str(x)[0:10])#将原有的时间格式变为字符串格式2
    left_11=pd.Series(['ID','机构','班级','姓名首拼','学生姓名','昵称','性别','上年课时结余','购买课时','目前剩余课时','上课数量统计汇总'])
    # new_title=left_6.append(df.iloc[0].str.cat(df.iloc[1],sep=',')[6:]).tolist() #构建新的表头，使用了函数 df.iloc[0].str.cat
    # title_time=left_11.append(df.iloc[0][11:].apply([0:8])).tolist()
    title_time=df.iloc[0,11:].fillna('-').apply(lambda x: x if x=='-' else x[0:8])

    title_time=left_11.append(df.iloc[0,11:].fillna('-').apply(lambda x: x if x=='-' else x[0:8])).tolist()

    title_crs=left_11.append(df.iloc[0,11:].fillna('-').apply(lambda x: x if x=='-' else x[9:])).tolist()
    

    #学生实际上的课表
    df_std=df.iloc[1:]
    df_std_new=df_std.copy()
    
    df_std_new.columns=title_time

    for i in range(0,df_std_new.shape[0]):
        for j in range(11,df_std_new.shape[1]):
            if df_std_new.iloc[i,j]=='√':
                df_std_new.iloc[i,j]=title_crs[11:][j-11]
    
    #总课表
    df_crs_0=df.iloc[0,11:].copy().T
    df_crs_0.fillna('-',inplace=True)
    df_crs=df_crs_0.apply(lambda x: x if x=='-' else x[0:8]).to_frame()
    df_crs.columns=['上课日期']
    df_crs['课程名称']=df_crs_0.apply(lambda x: x if x=='-' else x[9:])
    df_crs.reset_index(inplace=True,drop=True)
    df_crs.replace('-',np.nan,inplace=True)
    
    return {'total_crs':df_crs,'std_crs':df_std_new}

def std_term_crs(std_name='黄建乐',start_date='20000927',end_date='21000105',xls='D:\\Documents\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生信息表\\2021春-学生信息表（周六）.xlsx'):
    df=crs_sig_table(xls=xls)   
    std_df=df['std_crs']
    std_name=std_name.strip()
    # print(std_df)
    infos=std_df[std_df['学生姓名']==std_name]
    info_basic=infos[['机构','班级','姓名首拼','性别','ID','学生姓名','上课数量统计汇总']]
    info_crs_0=infos.iloc[:,11:]        
    info_crs=copy.copy(info_crs_0)
    info_crs.loc['aa']=info_crs_0.columns.values
    info_crs=info_crs.T
    info_crs.reset_index(drop=True,inplace=True)
    # print(info_crs)
    info_crs.columns=['课程名称','上课日期']    
    info_crs.replace('-',np.nan,inplace=True)
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

def std_feedback(std_name='韦宇浠',xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\2020秋-学生课堂学习情况反馈表（周二）.xlsx'):
    # wd='周'+days_calculate.num_to_ch(str(weekday))
    df_cmt=pd.read_excel(xls,sheet_name='课堂情况反馈表',skiprows=1)
    df_cmt.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名', \
                        'Unnamed: 5':'昵称','Unnamed: 6':'性别'},inplace=True)
    df_ability_total=pd.read_excel(xls,sheet_name='学员能力评分表',skiprows=1)
    df_ability_total.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名', \
                        'Unnamed: 5':'昵称','Unnamed: 6':'性别','Unnamed: 7':'优点特性','Unnamed: 8':'提升特性'},inplace=True)
    df_ability=df_ability_total[['学生姓名','理解力','空间想象力','逻辑思维','创造力','表达力','情绪力','人际力','自控力','适应力','学习力']]
    # print(df_ability)
    df_term_comment_txt=df_cmt.filter(regex='学期总结')
    df_term_comment=pd.concat([df_cmt[['ID','学生姓名']],df_term_comment_txt],axis=1)

    return {'df_ability':df_ability,'df_term_comment':df_term_comment}

def std_all_scores(xls_dir='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园',plus_tiyan='no'):
    xlsxs=[]
    for fn in os.listdir(os.path.join(xls_dir,'学生信息表')):
        if re.match(r'^\d.*-.*）.xlsx',fn):
            if plus_tiyan=='yes':
                xlsxs.append(os.path.join(xls_dir,'学生信息表',fn))
            else:
                if fn[-8:-6] != '体验':
                    xlsxs.append(os.path.join(xls_dir,'学生信息表',fn))

    df_infos=[]
    df_crss=[]
    df_acts=[]
    df_verifys=[]
    for xlsx in xlsxs:
        df_infos_pre=pd.read_excel(xlsx,sheet_name='学生档案表',skiprows=1,header=None,usecols=[0,1,2,3,4,5,6])
        df_crss_pre=pd.read_excel(xlsx,sheet_name='课堂积分',skiprows=2,header=None,usecols=[0,1,2,3,4,5,6,7])
        df_acts_pre=pd.read_excel(xlsx,sheet_name='活动积分',skiprows=2,header=None,usecols=[0,1,2,3,4,5,6,7])
        df_verifys_pre=pd.read_excel(xlsx,sheet_name='积分核销表',skiprows=1,header=None)
        df_infos.append(df_infos_pre)
        df_crss.append(df_crss_pre)
        df_acts.append(df_acts_pre)
        df_verifys.append(df_verifys_pre)

    df_info=pd.concat(df_infos)
    df_info.columns=['ID','机构','班级','姓名首拼','学生姓名','昵称','性别']
    df_info.dropna(how='all',axis=0,inplace=True)
    df_info=df_info[['ID','机构','姓名首拼','学生姓名','性别']]
    df_info.drop_duplicates('学生姓名',inplace=True)
    df_info.reset_index(inplace=True)

    df_crs=pd.concat(df_crss)
    df_crs.columns=['ID','机构','班级','姓名首拼','学生姓名','昵称','性别','课堂总积分']
    df_crs.dropna(how='all',axis=0,inplace=True)
    df_crs=df_crs[df_crs['学生姓名']!=0]
    # print(df_crs)

    df_act=pd.concat(df_acts)
    df_act.columns=['ID','机构','班级','姓名首拼','学生姓名','昵称','性别','活动总积分']
    df_act.dropna(how='all',axis=0,inplace=True)
    df_act=df_act[df_act['学生姓名']!=0]

    df_verify=pd.concat(df_verifys)
    df_verify.columns=['ID','机构','班级','姓名首拼','学生姓名','昵称','性别','核销积分','核销日期','兑换礼品','备注']
    df_verify.dropna(how='all',axis=0,inplace=True)
    std_verify_score=df_verify.groupby('学生姓名')['核销积分'].sum()
    
    std_verify_score=std_verify_score.to_frame()
    std_verify_score.reset_index(inplace=True)

    std_crs_score=df_crs.groupby('学生姓名')['课堂总积分'].sum()
    std_crs_score=std_crs_score.to_frame()
    std_crs_score.reset_index(inplace=True)

    std_act_score=df_act.groupby('学生姓名')['活动总积分'].sum()
    std_act_score=std_act_score.to_frame()
    std_act_score.reset_index(inplace=True)

    df_scores=pd.merge(df_info,std_crs_score,how='left',on='学生姓名')
    df_scores=pd.merge(df_scores,std_act_score,how='left',on='学生姓名')       
    df_scores['总积分']=df_scores.apply(lambda x:x['课堂总积分']+x['活动总积分'],axis=1)
    df_scores=pd.merge(df_scores,std_verify_score,how='left',on='学生姓名')
    df_scores.iloc[:,5:]=df_scores.iloc[:,5:].fillna(0)     
    df_scores['剩余积分']=df_scores.apply(lambda x:x['课堂总积分']-x['核销积分'],axis=1)
    df_scores=df_scores.iloc[:,1:]     
    

    # print(df_scores)
    return df_scores

def std_score_this_crs(xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生信息表\\2021秋-学生信息表（周一）.xlsx'):
    df=pd.read_excel(xls,sheet_name='课堂积分',header=None)
    res_date_crs=df.iloc[0,8:]   
    #将数字改为nan后去掉nan，只保留有课程名称的记录     
    date_crs=res_date_crs.apply(lambda x : np.nan if isinstance(x, int) else x) 
    date_crs.dropna(inplace=True)
    # print(date_crs)
    std_all_crs_scores=df.iloc[2:,:]
    remain_col_nums=[3+4*x for x in range(date_crs.shape[0])]
    std_scores=std_all_crs_scores.iloc[:,8:]
    std_scores=std_scores.iloc[:,remain_col_nums]
    std_info=std_all_crs_scores.iloc[:,:8]
    std_this_scores=pd.concat([std_info,std_scores],axis=1)

    col_names=['ID','机构','班级','姓名首拼','学生姓名','昵称','性别','课堂总积分']
    col_names.extend(date_crs.tolist())
    std_this_scores.columns=col_names
    # print(std_this_scores)
    #计算奖牌个数
    count_medal=df.iloc[2:,8:8+4*date_crs.shape[0]]
    for k,title in enumerate(date_crs.tolist()):        
        count_medal[title]=count_medal.iloc[:,4*k:4*(k+1)-1].sum(axis=1)

    medal_num=pd.concat([std_info,count_medal.iloc[:,4*date_crs.shape[0]:]],axis=1)
    medal_num.columns=col_names
    return {'std_this_scores':std_this_scores,'medals_this_class':medal_num}

def comments_after_class(cmt_date,crs_name_input,weekday,crs_list,std_info,tch_cmt):
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
    
    df_stdInfo=pd.read_excel(std_info,sheet_name='学生档案表')
    df_stdSig=pd.read_excel(std_info,sheet_name='学生上课签到表',skiprows=1)
                
    df_stdSig.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级', \
                                'Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名','Unnamed: 5':'昵称', \
                                    'Unnamed: 6':'性别','Unnamed: 7':'上期课时结余', \
                                        'Unnamed: 8':'购买课时','Unnamed: 9':'目前剩余课时','Unnamed: 10':'上课数量统计汇总'},inplace=True)
    # print(df_stdSig.columns)
    # print(df_stdSig.columns)
    Students_sig=df_stdSig.loc[df_stdSig[cmt_date+'-'+crs_code+crs_name].isin(['√','*'])][['机构','班级','姓名首拼','学生姓名']] #上课的学生名单
    # Students_sig=df_stdSig.loc[df_stdSig.iloc[:,11:]=='√'][['机构','班级','姓名首拼','学生姓名']] #上课的学生名单   
    # print(Students_sig)
    # print(Students_sig,df_stdInfo)         
    Students=pd.merge(Students_sig,df_stdInfo,on='学生姓名',how='left') #根据学生名单获取学生信息
    Students_List=Students.values.tolist()

    NumtoC={'1':'一','2':'二','3':'三','4':'四','5':'五','6':'六','7':'日'}
    # shtName='周'+NumtoC[str(weekday)]
    shtName='课堂情况反馈表'
    TeacherCmt=pd.read_excel(tch_cmt,sheet_name=shtName,skiprows=1)
    TeacherCmt.fillna('-',inplace=True)
    TeacherCmt.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼', \
                                'Unnamed: 4':'学生姓名','Unnamed: 5':'昵称','Unnamed: 6':'性别'},inplace=True)

    # print('完成')
    # print(Students_List)
    return {'std_list':Students_List,'crs_info':crs_info,'tch_cmt':TeacherCmt}    

# def term_summary_txt(std_name='韦宇浠',xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\学员课堂学习情况反馈表.xlsx',weekday=2):
#     wd='周'+days_calculate.num_to_ch(str(weekday))
#     df=pd.read_excel(xls,sheet_name=wd,skiprows=1)

def multi_std_infos(tb_dir='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\学生信息表'):
    fns=os.listdir(tb_dir)
    xlsxs=[]
    for fn in fns:
        if re.match(r'^\d.*-.*）.xlsx',fn):
            xlsxs.append(os.path.join(tb_dir,fn))
    
    # dfs.to_excel('e:/temp/kp.xlsx')
    # print(dfs)
    # print(df_ticks.iloc[3,0])
    # print(xlsxs,'\n')
    all_tables=[]
    for xlsx in xlsxs:        
        df_total=pd.read_excel(xlsx,sheet_name='学生上课签到表',header=None)
        df_basic=df_total.iloc[3:,0:7]
        df_ticks=df_total.iloc[3:,11:].copy()
        df_titles=df_total.iloc[2,11:]
        df_dates=df_total.iloc[1,11:]

        for i in range(0,df_ticks.shape[0]):
            for j in range(0,df_ticks.shape[1]):        
                    if df_ticks.iloc[i,j]=='√':
                        df_ticks.iloc[i,j]=df_titles.tolist()[j]
        df_ticks.columns=df_dates.tolist()
        df_basic.columns=['ID','机构','班级','姓名首拼','学生姓名','昵称','性别']
        # df_basic=df_basic.reset_index()
        dfs=pd.concat([df_basic,df_ticks],axis=1)
        dfs=dfs.reset_index(drop=True)
        dfs.dropna(how='all',axis=1,inplace=True)
        dfs.dropna(how='all',axis=0,inplace=True)
        all_tables.append(dfs)
    
    for k,tb in enumerate(all_tables):
        print(k,'\n',tb)

    # df_all=pd.concat(all_tables,ignore_index=True)
    # df_all.to_excel('e:/temp/kkkdd.xlsx')

def class_taken(xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生信息表\\2021春-学生信息表（周一）.xlsx'):
    df=pd.read_excel(xls,sheet_name='学生上课签到表',skiprows=1,header=0)
    cls_taken=df.iloc[:,11:].columns.values.tolist()
    df_std=df.iloc[:,4:5].values.tolist()
    std_names=[std[0] for std in df_std]

    return {'cls_taken':cls_taken,'std_names':std_names}


if __name__=='__main__':
    # print(std_feedback())
    # print(std_term_crs())
    # k=crs_sig_table()
    # print(k['total_crs'])
    # print(std_all_scores())
    # print(class_taken())
    print(std_score_this_crs())

    # crs_list="E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\2-乐高课程\\课程信息表.xlsx"
    # std_list="E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生信息表\\2021春-学生信息表（周一）.xlsx"
    # tch_cmt="E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\2021春-学生课堂学习情况反馈表（周一）.xlsx"
    # res=comments_after_class(cmt_date='20210329',crs_name_input='L040认识零件（二）',weekday=1,crs_list=crs_list,std_info=std_list,tch_cmt=tch_cmt)
    # print(res['tch_cmt'])

    # multi_std_infos()