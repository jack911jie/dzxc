import os
from subprocess import STD_ERROR_HANDLE
import sys
import copy
# import this
import pandas as pd
import numpy as np
from datetime import datetime
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.max_rows', None)


class StudentData:
    def __init__(self,wecomid='1688856932305542',place='01-超智幼儿园',template_fn='学生课堂行为评分标准表.xlsx'):
        self.work_dir='E:\\WXWork\\'+wecomid+'\\WeDrive\\大智小超科学实验室'
        self.place=place
        self.standard_dir=os.path.join(self.work_dir,place,'SOP')
        self.standard_fn=os.path.join(self.standard_dir,template_fn)

    def read_template(self):
        df_template=pd.read_excel(self.standard_fn,sheet_name='行为及评分总表')
        return df_template

    def std_mark(self,std_name='邓恩睿',tb_name='1',std_mark_fn='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\反馈表\\2022\\2022春-学生课堂行为记录表（周一）.xlsx'):
    # def std_mark(self,std_name='邓恩睿',tb_name='1',std_mark_fn='C:\\Users\\jack\\Desktop\\副本2022春-学生课堂行为记录表（周一）.xlsx'):
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

    def std_crs(self,std_name='邓恩睿',std_crs_fn='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生信息表\\2022\\2022春-学生信息表（周一）.xlsx'):
        df_all_std_crs=pd.read_excel(std_crs_fn,sheet_name='学生上课签到表',skiprows=1)
        df_all_std_crs.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名',
                                        'Unnamed: 5':'昵称','Unnamed: 6':'性别','Unnamed: 7':'上期课时结余','Unnamed: 8':'购买课时','Unnamed: 9':'目前剩余课时','Unnamed: 10':'上课数量统计汇总'}, 
                            inplace=True)
        df_std_crs=df_all_std_crs[df_all_std_crs['学生姓名']==std_name].iloc[:,11:]
        df_std_crs=df_std_crs.T.reset_index()
        df_std_crs.columns=['课程日期及名称','是否上课']
        df_std_crs.drop(df_std_crs[df_std_crs['是否上课']!='√'].index,inplace=True)
        return df_std_crs

    def num_to_weekday(self,num):
        trans_list={1:'一',2:'二',3:'三',4:'四',5:'五',6:'六',7:'日'}
        return trans_list[num]

    def std_all_crs(self,std_name='邓恩睿',in_list=[['2022春',1],['2021秋',1]],end_time=''):
        
        crs_info_dir=os.path.join(self.work_dir,self.place,'学生信息表')
        _df_all_crs=[]
        for tbl in in_list:
            crs_tbl_fn=os.path.join(crs_info_dir,tbl[0][:4],tbl[0]+'-学生信息表（周'+self.num_to_weekday(tbl[1])+'）.xlsx')
            one_crs=self.std_crs(std_name=std_name,std_crs_fn=crs_tbl_fn)
            _df_all_crs.append(one_crs)
        df_all_crs=pd.concat(_df_all_crs)
        df_all_crs['上课日期']=df_all_crs['课程日期及名称'].apply(lambda x:datetime.strptime(x[0:8],'%Y%m%d'))
        df_all_crs['课程名称']=df_all_crs['课程日期及名称'].str[9:]
        df_all_crs.sort_values(by=['上课日期'],inplace=True)        
        #如有截止时间
        if end_time!='':
            # print('end_time exists')
            df_all_crs=df_all_crs[df_all_crs['上课日期']<=datetime.strptime(end_time,'%Y%m%d')]

        df_all_crs.reset_index(drop=True,inplace=True)

        return df_all_crs

    def multi_tbl_score(self,std_name='邓恩睿',in_list=[['2022春',1]],end_time=''):
        #学生真实上课表
        std_real_crs=self.std_all_crs(std_name=std_name,in_list=in_list,end_time=end_time)

        
        
        #读取行为记录表的课程
        _all_beh=[]
        for tbl in in_list:
            beh_rec_fn=os.path.join(self.work_dir,self.place,'每周课程反馈','反馈表',tbl[0][0:4],tbl[0]+'-学生课堂行为记录表（周'+self.num_to_weekday(tbl[1])+'）.xlsx')
            
            df_beh_crs=pd.read_excel(beh_rec_fn,sheet_name='课程信息表',skiprows=1,usecols='a:e')

            df_beh_crs['学期']=tbl[0][0:5]
            df_beh_crs.drop(df_beh_crs[df_beh_crs['上课时间']=='-'].index,inplace=True)
            df_this_all=[]
            #如真实上过课，则读取课程的行为分数，并匹配上日期及课程
            for std_real_took in std_real_crs['课程日期及名称'].tolist():                
                tb_name=str(df_beh_crs[df_beh_crs['课程名称']==std_real_took]['课时'].tolist()[0])
                _this_crs_score=self.std_mark(std_name=std_name,tb_name=tb_name,std_mark_fn=beh_rec_fn)
                _this_crs_score['学期']=tbl[0][0:5]
                _this_crs_score['节次']=tb_name
                _this_crs_score['课程编码及名称']=std_real_took
                _this_crs_score['上课日期']=datetime.strptime(std_real_took[0:8],'%Y%m%d')
                df_this_all.append(_this_crs_score)

        this_std_all_score=pd.concat(df_this_all)

        # this_std_all_score.to_clipboard()
        return this_std_all_score

    def batch_deal_std_scores(self,std_list,terms,output_name):
        _dfs=[]
        for std in std_list:
            df=self.multi_tbl_score(std_name=std,in_list=terms)
            _dfs.append(df)

        df=pd.concat(_dfs)
        df.to_excel(output_name)

        print('完成')
        return df

    def batch_different_term(self,output_name,std_terms,end_time=''):
        std_term_mix=[]
        for term,stds in std_terms:
            for std in stds:
                std_term_mix.append([std,term[0]])
        
        _dfs=[]
        for number,std_term_name in enumerate(std_term_mix):          
            print('正在生成第 '+str(number+1)+'条（共 '+str(len(std_term_mix))+'条）-'+std_term_name[0]+ '的数据……',end='') 
            # print(std_term_name[0],std_term_name[1])
            df=self.multi_tbl_score(std_name=std_term_name[0],in_list=[std_term_name[1]],end_time=end_time)
            _dfs.append(df)
            print('完成')

        print('正在合并数据')
        df=pd.concat(_dfs)
        print('正在保存到本地')
        df.to_excel(output_name)
        os.startfile(os.path.dirname(output_name))

        print('完成')
        return df


if __name__=='__main__':  
    std_list1=['邓恩睿','邓立文','黄文俊','黄昱涵','李俊豪','廖世吉','李贤斌','磨治丞','农淑颖','农雨蒙','覃熙雅','陶梓翔','韦欣彤','韦欣怡','吴岳']
    terms1=[['2022春',1]]
    std_list2=['李崇析','陈锦媛','陆浩铭','唐欣语','邹维韬','朱端桦','谢威年','韦宇浠','韦启元','沈芩锐','岑亦鸿','廖茗睿','黄进桓','黄钰竣','韦万祎']
    terms2=[['2022春',5]]
    p=StudentData(wecomid='1688856932305542',place='001-超智幼儿园',template_fn='学生课堂行为评分标准表.xlsx')
    # res=p.multi_tbl_score(std_name='李崇析',in_list=[['2022春',5]],end_time='20220609')
    
    # res=p.batch_deal_std_scores(std_list=std_list,terms=terms,output_name='e:/temp/temp_dzxc/result.xlsx')
    # print(res)

    p.batch_different_term('e:/temp/temp_dzxc/result.xlsx',[[terms1,std_list1],[terms2,std_list2]],end_time='20220609')