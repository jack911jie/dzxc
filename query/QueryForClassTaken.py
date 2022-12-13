import os
import sys
from numpy.lib.arraysetops import isin
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import WashData
import readConfig
import days_calculate
import pandas as pd 
import re
import copy

class Query:
    def __init__(self,wecom_id='1688856932305542',place_input='001-超智幼儿园'):
        config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'query.config'))
        self.std_in_class_list=config['学生分班表']
        self.std_in_class_list=self.std_in_class_list.replace('$',wecom_id).replace('place',place_input)
        self.std_info_dir=config['机构文件夹']
        self.std_info_dir=self.std_info_dir.replace('$',wecom_id).replace('place',place_input)

    def std_class_taken(self,weekday=[6],display='print_list',format='only_clsname'):
        wds=[]
        for wkdy in weekday:
            wds.append(days_calculate.num_to_ch(wkdy))
        cls=[]
        fn_dir=os.path.join(self.std_info_dir,'学生信息表')
        for xls in os.listdir(fn_dir):
            if re.match(r'\d\d\d\d\w{1}-.*.xlsx$',xls) and xls[-8:-6]!='体验':
                for wd in wds:
                    if xls[-7:-6]==wd:
                        fn=os.path.join(fn_dir,xls)
                        info_taken=WashData.class_taken(fn)
                        cls.extend(info_taken['cls_taken'])
        cls_only_name=[item[9:] for item in cls]


        if format=='only_clsname':
            cls=cls_only_name
        
        if display=='print_list':
            for line in cls:
                print(line)
        
        return cls

    def deep_fn(self,term,dir,crs_name,std_list):
        # df_current=pd.read_excel(fn,skiprows=1)
        # std_list=[item[0] for item in df_current.iloc[:,4:5].values.tolist()]
        # print(std_list)
        rpt=0
        rpt_info=[]
        for root,dirs,fns in os.walk(dir):
            for xls in fns:
                # print('---------------------------------------------------')
                if re.match(r'\d\d\d\d\w{1}-.*.xlsx$',xls) and xls[-8:-6]!='体验':
                    fn_ever=os.path.join(root,xls)
                    info_taken=WashData.class_taken(fn_ever)
                    # print('info-taken std names:',info_taken['std_names'])
                    try:
                        cls_taken_get=[cls[9:] for cls in info_taken['cls_taken']]
                        #输入的课程名单是否在既往的课程中
                    except:
                        pass

                    # print('cls taken get:',cls_taken_get)
                    
                    if crs_name in cls_taken_get:
                        # print('now crsname:',crs_name)
                        #当前学期的学生姓名是否在过去的名单中
                        for std in std_list:
                            # print('std_name',std)
                            if std in info_taken['std_names']:
                                # print('in list std',std)
                                rpt_info.append([std,crs_name,xls[:5],xls[-8:-6]])
                                rpt+=1

        # print(rpt_info)
        return [rpt_info,rpt]
        

    def check_duplicate(self,term='2021秋',weekday=5,crs_name='L103摩托车骑士',show='show'):
        wd=days_calculate.num_to_ch(weekday)
        fn_dir=os.path.join(self.std_info_dir,'学生信息表')
        fn=os.path.join(fn_dir,term[:4],term+'-学生信息表（周'+wd+'）.xlsx')
        df_current=pd.read_excel(fn,skiprows=1)
        std_list=[item[0] for item in df_current.iloc[:,4:5].values.tolist()]

        info_current=[crs_name,std_list]
        # rpt=0
        # rpt_info=[]

        rpt_info,rpt=self.deep_fn(term=term,dir=fn_dir,crs_name=crs_name,std_list=std_list)
        
        # for root,dirs,kk in os.walk('E:\\temp\\temp_dzxc'):
        #     print(kk)
        #     if re.match(r'\d\d\d\d\w{1}-.*.xlsx$',xls) and xls[-8:-6]!='体验':
        #         fn_ever=os.path.join(fn_dir,xls)
        #         info_taken=WashData.class_taken(fn_ever)
        #         cls_taken_get=[cls[9:] for cls in info_taken['cls_taken']]
        #         #输入的课程名单是否在既往的课程中
                
        #         if crs_name in cls_taken_get:
        #             #当前学期的学生姓名是否在过去的名单中
        #             for std in std_list:
        #                 if std in info_taken['std_names']:
        #                     rpt_info.append([std,crs_name,xls[:5],xls[-8:-6]])
        #                     rpt+=1
        if show=='show':
            if rpt>0:
                print('\n在 {} 学期 星期{} 的学生课程 {} 中检测到以下重复\n'.format(term,wd,crs_name))
                for info in rpt_info:
                    print(info)
            else:
                print('未检测到重复值')

        return [rpt_info,rpt]

    def check_conflict(self,term='2021秋',weekday=5,fn='c:/Users/jack/desktop/w5待排课程.txt',show_res='yes',write_file='yes'):
        with open(fn,'r',encoding='utf-8') as f:
            lines=f.readlines()
        to_arrange=[itm.strip() for itm in lines]
        res=[]
        rpt=0
        for crs in to_arrange:
            res_dup=self.check_duplicate(term=term, weekday=weekday,crs_name=crs,show='no')
            rpt=rpt+res_dup[1]
            check_res=res_dup[0]
            _res=[]
            for _check_res in check_res:
                _res.append([_check_res[0],_check_res[2],_check_res[3]])
            res.append([crs,_res])
        
        # print(res)

        if show_res=='yes':
            if rpt>0:
                print('\n在 {} 学期 星期{} 的班级课程中检测到以下重复值：'.format(term,days_calculate.num_to_ch(weekday)))
                res_no_dups=[]
                for conf in res:
                    if conf[1]:
                        if conf[0]!='':
                            print('\n'+conf[0]+'---->')
                            for conff in conf[1]:
                                print(conff[0]+'  '+conff[1]+'  '+conff[2])
                    else:
                        if conf[0]!='':
                            res_no_dups.append(conf[0])
                if res_no_dups:
                    print('\n=========分==隔==线===========\n')
                    for res_no_dup in res_no_dups:
                        print(res_no_dup)
                    print('-------------------------------->\n在 {} 学期 星期{} 的班级课程中未检测到重复值。\n'.format(term,days_calculate.num_to_ch(weekday)))
                    print('\n')
            else:
                print('\n')
                for conf in res:
                    if conf[0]!='':
                        print(conf[0])
                print('----------------------------------------->')
                print('\n以上在 {} 学期 星期{} 的班级课程中未检测到重复值。\n'.format(term,days_calculate.num_to_ch(weekday)))

        if write_file=='yes':
            res_fn='c:/Users/jack/desktop/crs_conflict_result.txt'
            if rpt>0:
                with open (res_fn,'w',encoding='utf-8') as fw:
                    fw.write('\n在 {} 学期 星期{} 的班级课程中检测到以下重复值：\n'.format(term,days_calculate.num_to_ch(weekday)))
                    res_no_dup_w=[]
                    for conf in res:
                        if conf[1]:
                            fw.write('\n'+conf[0]+'---->'+'\n')
                            for conff in conf[1]:
                                fw.write(conff[0]+'  '+conff[1]+'  '+conff[2]+'\n')
                        else:
                            if conf[0]!='':
                                res_no_dup_w.append(conf[0])
                    if res_no_dup_w:
                        fw.write('\n=========分==隔==线===========\n\n')
                        for res_no_dup in res_no_dups:
                            fw.write(res_no_dup+'\n')
                        fw.write('------------------------------->\n在 {} 学期 星期{} 的班级课程中未检测到重复值。\n'.format(term,days_calculate.num_to_ch(weekday)))
                        fw.write('\n')
                    print('\n文件写入完成')
            else:
                with open (res_fn,'w',encoding='utf-8') as fw:
                    fw.write('\n')
                    for conf in res:
                        if conf[0]!='':
                            fw.write(conf[0]+'\n')
                    fw.write('----------------------------------->'+'\n')                
                    fw.write('\n在 {} 学期 星期{} 的班级课程中未检测到重复值。\n'.format(term,days_calculate.num_to_ch(weekday)))

        # print(res)
        return res

        # print(check_res)

        
    def check_conflict_std(self,term='2021秋',weekday=5,cls=1,fn='c:/Users/jack/desktop/w5待排课程.txt',show_mode='crs'):
        with open(fn,'r',encoding='utf-8') as f:
            lines=f.readlines()
        to_arrange=[itm.strip() for itm in lines]

        df_stds=pd.read_excel(os.path.join(self.std_info_dir,'学生信息表','学生分班表.xlsx'))
        _df_cls=df_stds[(df_stds['分班']=='w'+str(weekday)+str(cls).zfill(2)) & (df_stds['学期']==term)]
        df_cls=copy.deepcopy(_df_cls)

        if df_cls.shape[0]==0:
            print('\n在学生分班表中未查询到符合条件的学生名单')
            return
        else:
            df_cls['学生编码及姓名']=df_cls['ID']+df_cls['学生姓名']
            res_cross=[]
            res_cross_cls=[]
            for std in df_cls['学生编码及姓名'].tolist():
                # print('\n正在检查 {} 的课程……\n'.format(std))
                try:
                    df_std=pd.read_excel(os.path.join(self.std_info_dir,'学生档案',std+'.xlsx'),sheet_name='课程记录')
                    # print(df_std)
                    crs_taken=df_std['课程编码及名称'].tolist()
                    cross=list(set(to_arrange).intersection(set(crs_taken)))
                    if cross:
                        # print('{} 有以下课程重复：{} \n----------------------------------------\n'.format(std[6:],','.join(cross)))
                        res_cross.append([std,cross])
                        # cls_asp=[]
                        # for crs_cls in cross:
                        #     if crs_cls not in cls_asp:
                        #         cls_asp.append(crs_cls)
                except Exception as err:
                    print('错误：{}'.format(err))
        
        if len(res_cross)==0:
            print('\n未发现课程冲突')
            return 
        else:
            #课程去重
            crs_list=[]
            for itm in res_cross:
                crs_list.extend(itm[1]) 

            crs_lists=list(set(crs_list))

            crs_asp=[]
            for crs_name in crs_lists:
                std_names=[]
                for std_crs in res_cross:
                    if crs_name in std_crs[1]:
                        if std_crs[0] not in std_names:
                            std_names.append(std_crs[0])
                crs_asp.append([crs_name,std_names])

            ###显示
            if res_cross:
                if show_mode=='':
                    pass
                elif show_mode=='crs':
                    print('\n以下课程重复:')
                    for itm in crs_asp:
                        print('\n{} ------->\n{}\n-------------------'.format(itm[0],'，'.join(itm[1])))
                else:
                    print('\n以下学生有重复课程：')
                    for itm in res_cross:
                        print('\n{} ------->\n{}\n-------------------'.format(itm[0],'，'.join(itm[1])))

                

            #crs_asp:从重复课程角度显示学生姓名
            #res_cross:从学生姓名角度显示重复课程
            
            return {'crs_asp':crs_asp,'res_cross':res_cross}


class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value
        


if __name__=='__main__':
    qry=Query(place_input='001-超智幼儿园')
    res=qry.check_conflict_std(term='2022秋',weekday=5,cls=1,fn='c:/Users/jack/desktop/待排课程.txt',show_mode='crs')
  
    # qry.check_duplicate(term='2021秋',weekday=1,crs_name='L026跷跷板') 
    # conflict=qry.check_conflict(term='2022秋',weekday=5,fn='c:/Users/jack/desktop/待排课程.txt',show_res='yes',write_file='no')
    # qry.test()
