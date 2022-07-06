import os
from subprocess import STD_ERROR_HANDLE
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import WashData
import copy
import pypinyin
# import this
import re
import pandas as pd
import numpy as np
from datetime import datetime
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.max_rows', None)
# pd.set_option('display.max_columns', None)


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
        df_base_tbl=df_print_tbl[['环节','课堂项目','描述','打印编码','细分编码','分数']]
        # print(df_print_tbl)
        # print(df_base_tbl)

        def one_table(tb_name=tb_name):
            df_std_one=pd.read_excel(std_mark_fn,sheet_name=tb_name,skiprows=3,usecols='D:S')
            df_std_one.rename(columns={'Unnamed: 3':'行为描述'},inplace=True)
            df_std=df_std_one[['行为描述',std_name]]
            
            df_res=pd.concat([df_base_tbl,df_std],axis=1)
            _df_std_out=df_res[['打印编码','行为描述','细分编码','分数',std_name]]
            
            df_std_out=copy.deepcopy(_df_std_out)
            df_std_out.fillna(0,inplace=True)

            return df_std_out

        def one_tbl_score():
            this_std_score=one_table(tb_name=tb_name)

            
            #筛选出标记有分数的项
            marked_std_score=this_std_score.drop(this_std_score.loc[this_std_score[std_name]==0].index)

            
            ins_tbl=mid_score_tbl.merge(marked_std_score,how='left',on='细分编码')
            ins_tbl.to_clipboard()
            ins_tbl['行为描述']=ins_tbl.apply(lambda x: x['行为描述_x'] if pd.isna(x['行为描述_y']) else x['行为描述_y'],axis=1)
            ins_tbl['分数']=ins_tbl.apply(lambda x: x['分数_y'] if x[std_name]==1 else x['分数_x'],axis=1)   
            ins_tbl['姓名']=std_name         
            # ins_tbl['标准分']=ins_tbl['分数']*100/3
            

            #筛选出结果
            # print(ins_tbl)
            res_ins_tbl=ins_tbl[['姓名','环节','一级能力编码','一级能力名称','二级能力编码','二级能力名称','细分编码','细分内容','打印编码_x','行为描述','分数']]
            res_ins_tbl.columns=['姓名','环节','一级能力编码','一级能力名称','二级能力编码','二级能力名称','细分编码','细分内容','打印编码','行为描述','分数']

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

    def chr_to_caption(self,chr):
        all_py=''
        duoyin={'覃':1}
        for ss in chr:
            if ss in list(duoyin.keys()):
                py_pos=duoyin[ss]
            else:
                py_pos=0
            s_cap=pypinyin.pinyin(ss,heteronym=True,style=pypinyin.NORMAL)[0][py_pos][0].upper()
            all_py+=s_cap
        return all_py
            
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
                _this_crs_score['姓名首拼']=_this_crs_score['姓名'].apply(lambda x:self.chr_to_caption(x))
                _this_crs_score['学期']=tbl[0][0:5]
                _this_crs_score['节次']=tb_name
                _this_crs_score['课程编码及名称']=std_real_took
                _this_crs_score['上课日期']=datetime.strptime(std_real_took[0:8],'%Y%m%d')
                df_this_all.append(_this_crs_score)

        this_std_all_score=pd.concat(df_this_all)
        this_std_all_score=this_std_all_score[['姓名首拼','姓名','环节','一级能力编码','一级能力名称','二级能力编码','二级能力名称','细分编码','细分内容','打印编码','行为描述','分数','学期','节次','课程编码及名称','上课日期']]
        this_std_all_score.columns=['姓名首拼','姓名','环节','一级指标编码','一级指标名称','二级指标编码','二级指标名称','三级指标编码','三级指标名称','四级指标编码','四级指标名称','分数','学期','节次','课程编码及名称','上课日期']
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


    # def pre_deal_score(self,terms=[['2022春',1]]):
    #     for k,term in enumerate(terms):
    #         info_xls=os.path.join(self.work_dir,self.place,'学生信息表',term[0][:4],term[0]+'-学生信息表（周'+self.num_to_weekday(term[1])+'）.xlsx')
    #         this_info=WashData.std_score_this_crs(xls=info_xls)
    #         _score=this_info['std_this_scores']
    #         _medals=this_info['medals_this_class']
    #         # scores.append(_score)
    #         if k>0:
    #             score=pd.merge(score,_score,how='outer',on='学生姓名')                
    #             medals=pd.merge(medals,_medals,how='outer')
    #         else:
    #             score=_score
    #             medals=_medals

    #     return {'score':score,'medals':medals}


    def teacher_cmt(self,crs_list,std_name='吴岳',terms=[['2022春',1]],term_cmt_date='20220704'):
        
        std_mark=[]
        for k,term in enumerate(terms):
            std_info_fn=os.path.join(self.work_dir,self.place,'学生信息表',term[0][:4],term[0]+'-学生信息表（周'+self.num_to_weekday(term[1])+'）.xlsx')
            cmt_fn=os.path.join(self.work_dir,self.place,'每周课程反馈','反馈表',term[0][:4],term[0]+'-学生课堂学习情况反馈表（周'+self.num_to_weekday(term[1])+'）.xlsx')
            std_mark_fn=os.path.join(self.work_dir,self.place,'每周课程反馈','反馈表',term[0][:4],term[0]+'-学生课堂行为记录表（周'+self.num_to_weekday(term[1])+'）.xlsx')

            try:
                _std_mark=pd.read_excel(std_mark_fn,sheet_name='课程信息表',skiprows=1,usecols='A:E')
                std_mark.append(_std_mark)
            except Exception as e:
                print(e)

        #将学生信息表内的积分拼接
            this_info=WashData.std_score_this_crs(xls=std_info_fn)
            _score=this_info['std_this_scores']
            _medals=this_info['medals_this_class']
            
        #将terms内的评论表拼接
            _df_cmt=pd.read_excel(cmt_fn,skiprows=1)
            _df_cmt.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名','Unnamed: 5':'昵称','Unnamed: 6':'性别'},inplace=True)
            clmns=[]
            smrys=[]
            clmns_term={}
            smrys_term={}
            # print(list(_df_cmt.columns))
            k_n=0
            for clmn_name in list(_df_cmt.columns):                
                if re.match(r'^\d{8}-L\d{3}.*',str(clmn_name)):
                    clmns.append(clmn_name)
                    clmns_term[clmn_name]=[term[0],str(k_n+1)]
                    k_n+=1

                if '能力' in str(clmn_name) or '心理' in str(clmn_name):
                    smrys.append(clmn_name)
                    smrys_term[clmn_name]=[term[0],0]

            for clmn in clmns:
                _df_cmt[clmn].fillna(_df_cmt[_df_cmt['学生姓名']=='通用评论'][clmn].tolist()[0],inplace=True)

            _df_cmt.drop(_df_cmt[_df_cmt['学生姓名']=='通用评论'].index,inplace=True)
            

            if k>0:
                df_cmt=pd.merge(df_cmt,_df_cmt,how='outer',on='学生姓名')
                df_score=pd.merge(df_score,_score,how='outer',on='学生姓名')
                df_medals=pd.merge(df_medals,_medals,how='outer',on='学生姓名')
                
            else:
                df_cmt=_df_cmt
                df_score=_score
                df_medals=_medals

        df_mark=pd.concat(std_mark)

        std_all_cmt=df_cmt[df_cmt['学生姓名']==std_name]

        # print(clmns_term)

        def replace_chr(txt,stdname,crsdatename):
            txt=txt.replace('#',stdname)
            txt_rplc=str(int(df_medals[df_medals['学生姓名']==std_name][crsdatename].tolist()[0]))+'枚积分币，共计'+str(int(df_score[df_score['学生姓名']==std_name][crsdatename].tolist()[0]))
            txt=txt.replace('*',txt_rplc)
            
            return txt

        crs_list['姓名']=std_name
        crs_list['姓名首拼']=crs_list['姓名'].apply(lambda x: self.chr_to_caption(x))
        crs_list['评论']=crs_list['课程日期及名称'].apply(lambda x: replace_chr(std_all_cmt[x].tolist()[0],std_name,x))
        crs_list['类型']='课后反馈'
        crs_list['学期']=crs_list['课程日期及名称'].apply(lambda x: clmns_term[x][0])
        crs_list['节次']=crs_list['课程日期及名称'].apply(lambda x: int(clmns_term[x][1]))
        crs_list['老师']=crs_list['课程日期及名称'].apply(lambda x: df_mark[df_mark['课程名称']==x]['主教老师'].tolist()[0])
        


        for smry in smrys:

            crs_list=crs_list.append({'课程日期及名称':smry[-8:]+'-'+smry[:-9],'是否上课':'√','上课日期':datetime.strptime(smry[-8:]+'000000','%Y%m%d%H%M%S'),'课程名称':smry[:-9],'姓名':std_name,
                              '姓名首拼':self.chr_to_caption(std_name),'评论':std_all_cmt[smry].tolist()[0],'类型':'学期评语','学期':smrys_term[smry][0],'节次':smrys_term[smry][1]},ignore_index=True)   



        return crs_list



    def batch_tch_cmt(self,output_name,std_terms,term_cmt_date='20220704',end_time=''):
        ress=[]
        _std_counts=0
        for terms,std_list in std_terms:
            _std_counts+=len(std_list)

        std_counts=0
        for terms,std_list in std_terms:
            for std_name in std_list:
                print('\n正在处理 ',std_name,'……（第',str(std_counts+1),'个/共',str(_std_counts),'个)',end='')
            # pre_res=self.pre_deal_score(terms=terms)       
                crs_list=self.std_all_crs(std_name=std_name,in_list=terms,end_time=end_time)
                res=self.teacher_cmt(crs_list=crs_list ,std_name=std_name,terms=terms,term_cmt_date=term_cmt_date)
                ress.append(res)
                print('完成')
                std_counts+=1
        print('\n正在合并数据')
        result=pd.concat(ress)
        print('\n正在保存数据至excel')
        result.to_excel(output_name)
        os.startfile(os.path.dirname(output_name))
        print('\n完成')

        return result



if __name__=='__main__':  
    std_list1=['邓恩睿','邓立文','黄文俊','黄昱涵','李俊豪','廖世吉','李贤斌','磨治丞','农淑颖','农雨蒙','覃熙雅','陶梓翔','韦欣彤','韦欣怡','吴岳']
    terms1=[['2022春',1]]
    std_list2=['李崇析','陈锦媛','陆浩铭','唐欣语','邹维韬','朱端桦','谢威年','韦宇浠','韦启元','沈芩锐','岑亦鸿','廖茗睿','黄进桓','黄钰竣','韦万祎']
    terms2=[['2022春',5]]
    p=StudentData(wecomid='1688856932305542',place='001-超智幼儿园',template_fn='学生课堂行为评分标准表.xlsx')
    # res=p.multi_tbl_score(std_name='李贤斌',in_list=[['2022春',1]],end_time='20220614')

    # res=p.std_mark(std_name='李贤斌',tb_name='2',std_mark_fn='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\反馈表\\2022\\2022春-学生课堂行为记录表（周一）.xlsx')

    # res.to_clipboard()

    # res=p.std_all_crs(std_name='吴岳',in_list=[['2022春',1]],end_time='')
    # print(res)
    # p.teacher_cmt(std_name='邓恩睿',terms=[['2022春',1],['2021秋',1]])

    res=p.batch_tch_cmt(output_name='e:/temp/temp_dzxc/result_cmt.xlsx',std_terms=[[terms1,std_list1],[terms2,std_list2]])


    # res.to_clipboard()
    # kk=p.chr_to_caption('邓恩睿')
    # print(kk)
    # res=p.batch_deal_std_scores(std_list=std_list,terms=terms,output_name='e:/temp/temp_dzxc/result.xlsx')
    # print(res)

    # p.batch_different_term('e:/temp/temp_dzxc/result.xlsx',[[terms1,std_list1],[terms2,std_list2]],end_time='20220714')