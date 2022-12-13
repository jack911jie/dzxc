import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import write_data
import WashData
import pandas as pd
import openpyxl
import re
import shutil
import pypinyin
import json
import numpy as np

class StudentClass:
    def __init__(self,place='001-超智幼儿园',wecom_id='1688856932305542',term='2022秋',weekday=5):
        self.place=place
        self.term=term
        self.wecom_id=wecom_id
        self.weekday=weekday

        with open(os.path.join(os.path.dirname(__file__),'config','class_record.config'),'r',encoding='utf-8') as f:
            lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=json.loads(_line)
        self.root_dir=config['机构根目录']
        self.root_dir=self.root_dir.replace('$',wecom_id).replace('place',place)

        self.wk_to_str={'1':'一','2':'二','3':'三','4':'四一','5':'五','6':'六','7':'日'}
        self.xls_info=os.path.join(self.root_dir,'学生信息表',term[:4],term+'-学生信息表（周'+self.wk_to_str[str(weekday)]+'）.xlsx')
        self.crs_type={'L':'乐高'}
        self.std_dis_table=os.path.join(self.root_dir,'学生信息表',)
        self.df_score=WashData.std_score_this_crs(xls=self.xls_info)
        self.df_sig=WashData.crs_sig_table(xls=self.xls_info)['std_crs']
        
        # print(self.df_sig)

    def read_sig(self,xls='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生信息表\\2022\\2022秋-学生信息表（周六）.xlsx'):
        df_sig=WashData.crs_sig_table(xls=xls)
        df_score=WashData.std_score_this_crs(xls=xls)
        # print(df_score,df_sig)
        return df_sig
                                                                                                                                                                                                                                                                      
                                                                                                
    def read_cmt(self,std_id_name='DZ0001黄建乐',crs_date_name='20221022-L175喂食的小鸟'):
       
        xls_cmt=os.path.join('E:\\WXWork',self.wecom_id,'WeDrive\\大智小超科学实验室',self.place,'每周课程反馈\\反馈表',self.term[:4],self.term+'-学生课堂学习情况反馈表（周'+self.wk_to_str[str(self.weekday)]+'）.xlsx')

        cmt=WashData.std_each_class_cmt(df_score=self.df_score,df_sig=self.df_sig,std_name=std_id_name[6:],crs_name=crs_date_name,
                        xls_cmt=xls_cmt)
        
        return cmt

    def append_cmt(self,std_id_name='DZ0001黄建乐',crs_date_name='20221022-L175喂食的小鸟',class_type='正式',tg_dir='E:\\temp\\temp_dzxc\\make_cus\\学生档案'):
        cmt=self.read_cmt(std_id_name=std_id_name,crs_date_name=crs_date_name)
        tg_xls=os.path.join(tg_dir,std_id_name+'.xlsx')

        df_tg=pd.read_excel(tg_xls,sheet_name='课程记录')
        df_tg['日期及课程编码名称']=df_tg['上课日期'].map(lambda x: str(x))+'-'+df_tg['课程编码及名称']

        try:
            txt_cmt=cmt['txt_cmt']
            score=cmt['score']

            if crs_date_name not in df_tg['日期及课程编码名称'].tolist():
                input_dataframe=pd.DataFrame(data=[[crs_date_name[:8],self.term,self.crs_type[crs_date_name[9].upper()],class_type,crs_date_name[9:],txt_cmt,score]],
                                                columns=['上课日期','学期','课程类型','上课类型','课程编码及名称','课程反馈','课堂积分'])
                write_log=write_data.WriteData().write_to_xlsx(input_dataframe=input_dataframe,output_xlsx=tg_xls,sheet_name='课程记录')
                print('{} {}'.format(crs_date_name,write_log))
            else:
                print('{} 已有记录'.format(crs_date_name))
        except Exception as err:
            print('错误：',err)

    def append_verified_score(self,std_id_name='DZ0001黄建乐',vfy_date=20221203):    
        df_vfys=pd.read_excel(self.xls_info,sheet_name='积分核销表')
        df_vfy=df_vfys[(df_vfys['核销日期']==int(vfy_date)) & (df_vfys['学生姓名']==std_id_name[6:]) & (df_vfys['ID']==std_id_name[:6])]        

        if df_vfy.shape[0]==0:
            print('{} 无 {} 在 {} 的积分兑换数据。'.format(self.xls_info.split('\\')[-1],std_id_name,str(vfy_date)))
            return np.nan
        else:
            df_write=df_vfy[['核销日期','核销积分','兑换礼品']]
            df_tg=pd.read_excel(os.path.join(self.root_dir,'学生档案',std_id_name+'.xlsx'),sheet_name='积分兑换')
            if int(vfy_date) not in df_tg['兑换日期'].tolist():
                print('\n正在尝试写入 {} {} 的兑换积分信息…'.format(std_id_name,str(vfy_date)),end='')
                write_basic_log=write_data.WriteData().write_to_xlsx(input_dataframe=df_write,output_xlsx=os.path.join(self.root_dir,'学生档案',std_id_name+'.xlsx'),sheet_name='积分兑换')
                print('\n',write_basic_log,end='')
                print('完成')       
            else:
               print('\n {} {} 的兑换积分信息已有记录……'.format(std_id_name,str(vfy_date)))

        


    def new_xlsx(self,std_name,sex='女',birthday='',tg_dir='E:\\temp\\temp_dzxc\\make_cus\\学生档案'):
        num=0
        for fn in os.listdir(tg_dir):
            if fn!='DZ0000学生档案模板.xlsx' and re.match(r'DZ\d{4}.*.xlsx',fn):
                num+=1

        new_fn=os.path.join(tg_dir,'DZ'+str(num+1).zfill(4)+std_name+'.xlsx')
        template_fn=os.path.join(tg_dir,'DZ0000学生档案模板.xlsx')
        print('正在生成 {} 档案，档案名：{} ……'.format(std_name,new_fn),end='')
        shutil.copyfile(template_fn,new_fn)        
        print('完成')

        self.write_basic_info(std_name=std_name,sex=sex,birthday=birthday,tg_xlsx=new_fn)

    def write_basic_info(self,std_name,sex,birthday,tg_xlsx):
        # id='DZ'+str(num+1).zfill(4)
        id=tg_xlsx.split('\\')[-1][:6]
        py=self.chr_to_caption(std_name)
        nickname=std_name[1:]
        basic_info=pd.DataFrame(data=[[id,py,std_name,nickname,sex,birthday]],columns=['ID','姓名首拼','姓名','昵称','性别','出生年月'])
        print('正在写入基本信息……',end='')
        write_basic_log=write_data.WriteData().write_to_xlsx(input_dataframe=basic_info,output_xlsx=tg_xlsx,sheet_name='基本情况')
        print('完成')

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

    def std_dis_table_batch(self,crs_date_name,weekday,cls,vfry_score='no'):
        tg_dir=os.path.join(self.root_dir,'学生档案')
        fn_std_class=os.path.join(self.root_dir,'学生信息表','学生分班表.xlsx')
        df_std_class=pd.read_excel(fn_std_class,sheet_name='分班表')
        df_std_class['学生编码及姓名']=df_std_class['ID']+df_std_class['学生姓名']
        std_list=df_std_class[df_std_class['分班']=='w'+str(weekday)+str(cls).zfill(2)]['学生编码及姓名']
        for std  in std_list:
            try:
                print('\n正在将课程评论及积分写入 {} 的个人档案中……'.format(std),end='')
                self.append_cmt(std_id_name=std,crs_date_name=crs_date_name,tg_dir=tg_dir)
                print('完成')

                if vfry_score=='yes':
                    print('\n查看积分兑换信息……'.format(std),end='')
                    self.append_verified_score(std_id_name=std,vfy_date=crs_date_name[:8])
                    print('完成')
            except Exception as err_batch:
                print('执行出错，错误代码：',err_batch)
            


if  __name__=='__main__':
    p=StudentClass(place='001-超智幼儿园',wecom_id='1688856932305542',term='2022秋',weekday=6)
    p.append_verified_score(std_id_name='DZ0028杨涵宇',vfy_date=20221203)
    # p.read_sig()          
    # res=p.read_cmt(std_id_name='DZ0001黄建乐',crs_date_name='20221022-L175喂食的小鸟')                                                      
    # print(res)
    # crs_lst=['20220903-L171公共汽车','20220917-L172自动巡逻机器人','20220924-L173扫雷车','20221015-L174机警的蚂蚱','20221022-L175喂食的小鸟','20221029-L176自动导弹发射车','20221105-L177自动售货机','20221112-L178避障赛车','20221119-L179人形机器人','20221126-L180B2轰炸机','20221203-L036毛毛虫']
    # std_lst=['DZ0034顾业熙','DZ0033刘泓彬','DZ0035李俊豪','DZ0054黄楚恒','DZ0051廖茗睿','DZ0032磨治丞','DZ0055刘晨凯','DZ0056陆一然','DZ0057潘子怡','DZ0058罗彬城']
    # std_lst=['DZ0028杨涵宇']
    # for std in std_lst:
    #     for crs in crs_lst:
    #         p.append_cmt(std_id_name=std,crs_date_name=crs,tg_dir='E:\\WXWork\\1688856932305542\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\学生档案')
    #         p.append_cmt(std_id_name=std,crs_date_name=crs,tg_dir='E:\\temp\\temp_dzxc\\make_cus\\学生档案')

    # new_list=[['黄楚恒','男',''],['刘晨凯','男',''],['陆一然','女',''],['潘子怡','女',''],['罗彬城','女','']]
    # for ln in new_list:
    #     p.new_xlsx(std_name=ln[0],sex=ln[1],birthday='',tg_dir='E:\\temp\\temp_dzxc\\make_cus\\学生档案')

