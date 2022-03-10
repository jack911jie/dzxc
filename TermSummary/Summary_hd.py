import os
import sys
# print(os.path.dirname(os.path.dirname(os.path.realpath(__file__))))
# sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'module'))
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),'modules'))
from composing import TxtFormat
from pic_transfer import Plot
import days_calculate
from readConfig import readConfig
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import WashData
import pandas as pd 
import matplotlib.pyplot as plt
from matplotlib.pyplot import MultipleLocator
import numpy as np
from PIL import Image,ImageDraw,ImageFont,ImageEnhance

class data_summary:
    def __init__(self):
        config=readConfig(os.path.join(os.path.dirname(os.path.realpath(__file__)),'configs','TermSummary.config'))
        self.std_dir=config['学生信息文件夹']
        self.design_dir=config['图纸文件夹']
        self.feedback_dir=config['学生课程反馈表文件夹']
        self.public_dir=config['公共素材文件夹']
        self.term_pic_dir=config['学员阶段总结文件夹']
        self.tch_sig_pic_dir=config['老师签名图片文件夹']

        font_cfg=readConfig(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'configs','dzxc_fonts.config'))
        self.font_list=font_cfg['fontList']
        self.font=TxtFormat().fonts
        
    def get_std_term_crs(self,std_name='韦宇浠',tb_list=[['2020秋','w2'],['2021春','w4']],start_date='20200901',end_date='20210521'):

        print('    正在读取 {} 的数据……'.format(std_name),end='')
        std_name=std_name.strip()
        std_term_crs=[]
        for tb in tb_list:
            term,weekday=tb[0],tb[1][-1]
            if tb[1][0].lower()=='w':
                xls_name=os.path.join(self.std_dir,'学生信息表',term[0:4],term+'-学生信息表（周'+days_calculate.num_to_ch(weekday)+'）.xlsx') 
            else:
                xls_name=os.path.join(self.std_dir,'学生信息表',term[0:4],term+'-学生信息表（'+std_name+'）.xlsx') 
                
            std_term_crs_pre=WashData.std_term_crs(std_name=std_name,start_date=start_date,end_date=end_date,xls=xls_name)
            std_term_crs.append(std_term_crs_pre)
            # print(std_term_crs_pre)
 
        df_std_crss=[]
        df_total_crss=[]
        for df_std_term_crs in std_term_crs:
            df_std_crss.append(df_std_term_crs['std_crs'])  
            df_total_crss.append(df_std_term_crs['total_crs'])
        
        df_std_crs=pd.concat(df_std_crss)
        df_std_crs.dropna(axis=0,inplace=True)
        df_std_crs.reset_index(inplace=True)
        df_total_crs=pd.concat(df_total_crss,ignore_index=True)

        #将不符合L001类型的数据删除，剔除请假补课等情况。
        df_std_crs.drop(df_std_crs[~(df_std_crs['课程名称'].str.contains(r'L\d{3}',regex=True))].index,inplace=True)

        # print(df_std_crs,'\n',df_total_crs)
        # print(df_std_crs)

        res_std_term_crs={'total_crs':df_total_crs,'std_crs':df_std_crs,'std_info':std_term_crs[0]['std_info']}
        print('完成')
        return res_std_term_crs

    def exp_poster(self,std_name='韦宇浠',start_date='20200930',end_date='20210121',cmt_date='20210407', \
                    tb_list=[['2020秋','w2'],['2021春','w4']],tch_name='阿晓老师',mode='all',k=1):
        print('正在处理……')
        info=self.get_std_term_crs(std_name=std_name,tb_list=tb_list,start_date=start_date,end_date=end_date)
        total_crs=info['total_crs']
        total_crs.dropna(inplace=True)
        std_crs=info['std_crs']
        std_crs.dropna(inplace=True)   
        std_info=info['std_info']

        #学期课程总数
        total_crs_num=total_crs.shape[0] 
        #格子的行数
        if total_crs_num/4==total_crs_num//4: 
            crs_lines=total_crs_num//4 
        else:
            crs_lines=total_crs_num//4+1

        #所有课程格子的总高度
        ht_crs=int((170*k+14*k)*crs_lines)

        #学生基本信息
        std_school=std_info['机构'].values[0]
        std_class=std_info['班级'].values[0]

        #学生学期末评语
        term,weekday=tb_list[-1][0],tb_list[-1][1][1]
        wd=days_calculate.num_to_ch(weekday)
        xls=os.path.join(self.feedback_dir,term+'-学生课堂学习情况反馈表（周'+wd+'）.xlsx')
        comments=WashData.std_feedback(std_name=std_name,xls=xls)['df_term_comment']    
        comments_for_std=comments[comments['学生姓名']==std_name][term+'学期总结-'+str(cmt_date)].values.tolist()[0].replace('#',std_name)
        std_crs_num=std_crs.shape[0]
        if  mode=='part':
            if  total_crs_num==std_crs_num:            
                std_crs_num_txt='{0}同学在上一阶段的{1} 节科学机器人课中，完成了全部课程的学习！'.format(std_name,total_crs_num)
            else:
                std_crs_num_txt='{0}同学在上一阶段的{1} 节科学机器人课中，完成了{2} 节课的学习，请假{3} 节。'.format(std_name,total_crs_num,std_crs_num,total_crs_num-std_crs_num)
        elif mode=='all':
            # if  total_crs_num==std_crs_num:            
            std_crs_num_txt='{0}已经完成了{1} 节科学机器人课程的学习！'.format(std_name,total_crs_num)
            # else:
                # std_crs_num_txt='{0}同学在上一阶段的{1} 节科学机器人课中，完成了{2} 节课的学习，请假{3} 节。'.format(std_name,total_crs_num,std_crs_num,total_crs_num-std_crs_num)
      
        comments_for_std=std_crs_num_txt+'\n'+comments_for_std
        font_size_cmt=int(30*k)
        prgh_nums=TxtFormat.split_txt_Chn_eng(wid=int(636*k),font_size=font_size_cmt,txt_input=comments_for_std)[1]
        #评语标题高度
        ht_cmt_title=int(60*k)
        #评语高度
        ht_prgh=int(font_size_cmt*1.5*(prgh_nums+4))

        

        ht_title=int(110*k)
        ht_std=int(100*k)
        ht_bottom=int(210)
        gap_0=int(12*k)
        gap_1=int(13*k)

        y_std_1=int(120*k)
        y_std_2=int((y_std_1+100))
        

        blocks=[ht_title,ht_std,ht_crs,ht_cmt_title,ht_prgh,ht_bottom]

        #计算总的高度
        ht_total_bg=int((sum(blocks)+gap_0*4+gap_1*(prgh_nums-1)))

        #课程标题坐标
        y_crs_title=int((y_std_2+gap_0+80))

        #课程格最左上角坐标        
        y_left_up=int((y_crs_title+gap_0))

        #评语绿底右下坐标
        y_cmt_bg=y_left_up+ht_crs+gap_0+ht_cmt_title+ht_prgh+30
        # draw.rectangle([(int(36*k),y_left_up+gap_0+ht_cmt_title+10),(int(684*k),y_cmt_bg)],fill='#F3F3E4')

        def draw_bg():
            bg=Image.new('RGBA',(int(720*k),ht_total_bg),'#F3F9E4')
            draw=ImageDraw.Draw(bg)
            #白底
            draw.rectangle([(int(13*k),int(110*k)),(int(707*k),ht_total_bg-25)],fill='#FFFFFF')
            #名字绿底
            draw.rectangle([(int(36),y_std_1),(int(250*k),y_std_2)],fill='#F3F9E4')
            draw.rectangle([(int(36),y_std_2),(int(684*k),y_std_2+2)],fill='#F3F9E4')            
            
            return bg

        def crs_pic(total_crs=total_crs,std_crs=std_crs,odr=0):
            wid_crs,ht_crs=int(150*k),int(170*k)

            #学生上过的课程和总课程的差集
            sub_set=total_crs.copy()
            sub_set=sub_set.append(std_crs)
            sub_set=sub_set.append(std_crs)
            sub_set.drop_duplicates(subset=['课程名称'],keep=False,inplace=True)

            crs_bg=Image.new('RGB',(wid_crs,ht_crs),'#F3F3E4')
            crs_name=total_crs.iloc[odr,:]['课程名称']
            crs_date=total_crs.iloc[odr,:]['上课日期']
            crs_pic=Image.open(os.path.join(self.design_dir,crs_name,crs_name[4:]+'.jpg'))
            crs_pic=crs_pic.crop((560,0,1280,720)).resize((int(110*k),int(110*k)))
            crs_bg.paste(crs_pic,(int(20*k),int(6*k)))
            draw=ImageDraw.Draw(crs_bg)
            ft_size=int(19*k)
            x_txt=int(((150*k-len(crs_name[4:])*ft_size)/2+3))
            
            draw.text((x_txt,int(122*k)),crs_name[4:],fill='#8fc31f',font=self.font('方正韵动粗黑简',ft_size)) #课程名称

            #判断是否有不上的课，如有，变灰
            if crs_name not in sub_set['课程名称'].values:
                draw.text((int(35*k),int(148*k)),str(crs_date)[:11],fill='#8fc31f',font=self.font('方正韵动粗黑简',int(12*k))) #课程日期
            else:
                draw.text((int(35*k),int(148*k)),'---------',fill='#8fc31f',font=self.font('方正韵动粗黑简',int(12*k))) #课程日期
                
                crs_bg=ImageEnhance.Brightness(crs_bg).enhance(1.4)
                crs_bg=crs_bg.convert('L')
            
            # crs_bg.show()
            return crs_bg

        def draw_crs():
            print('正在生成图片……',end='')

            #使用并改变母函数的变量
            nonlocal y_left_up
            bg=draw_bg()
            wid_crs,ht_crs=int(150*k),int(170*k)            

            draw=ImageDraw.Draw(bg)
            #分隔线
            draw.rectangle([(int(36*k),y_crs_title),(int(684*k),y_crs_title+2*k)],fill='#F3F3E4')

            cnt=0
            for i in range(crs_lines):
                x_begin=int(40*k)              
                for j in range(4):       
                    if cnt<total_crs_num:      
                        crs_block=crs_pic(total_crs=total_crs,std_crs=std_crs,odr=cnt)       
                        bg.paste(crs_block,(x_begin,y_left_up))
                        x_begin=x_begin+wid_crs+int(13*k)   
                        cnt+=1         
                y_left_up=y_left_up+ht_crs+gap_1 
            
            
            #分隔线
            draw.rectangle([(int(36*k),y_left_up+gap_0),(int(684*k),y_left_up+gap_0+2)],fill='#F3F3E4')
           #评语绿底
            # y_cmt_bg=y_left_up+gap_0+ht_cmt_title+ht_prgh
            draw.rectangle([(int(36*k),y_left_up+gap_0+ht_cmt_title+10),(int(684*k),y_cmt_bg)],fill='#F3F3E4')

            #二维码，logo
            _pic_logo=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超新logo.png').convert('RGBA')
            pic_logo=_pic_logo.resize((int(210*k),int(210*k/2.76)))
            r,g,b,a=pic_logo.split()

            y_logo=int((ht_total_bg-140))

            _pic_qrcode=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超视频号二维码2.png')
            pic_qrcode=_pic_qrcode.resize((int(100*k),int(100*k)))

            bg.paste(pic_logo,(int(40*k),int(y_logo+10)),mask=a)
            bg.paste(pic_qrcode,(int(580*k),int(y_logo-10)))

            #文字部分
            #标题
            draw.text((int(130*k),int(5*k)),'科学机器人课',fill='#8fc31f',font=self.font('汉仪超级战甲',int(78*k))) 
            #姓名班级
            draw.text((int(60*k*0.9),int(142*k*0.9)),std_name,fill='#F8B62D',font=self.font('汉仪字酷堂义山楷w',int(60*k))) 
            draw.text((int(400*k*0.9),int(182*k*0.9)),std_school+' '+std_class,fill='#6C8466',font=self.font('汉仪字酷堂义山楷w',int(30*k))) 
            draw.text((int(45),int((y_crs_title-52))),'点亮的课程',fill='#8fc31f',font=self.font('鸿蒙印品',int(35*k))) 
            draw.text((int(45),int((y_left_up+gap_0+20))),'学期评语',fill='#8fc31f',font=self.font('鸿蒙印品',int(35*k))) 

            #评语
            TxtFormat.put_txt_img(draw=draw,tt=comments_for_std,total_dis=int(628*k*0.95), \
                                  xy=[int(58),int((y_left_up+gap_0+110))],dis_line=int(20*k),fill='#595757', \
                                  font_name='丁永康硬笔楷书',font_size=font_size_cmt,addSPC="yes")

            draw.text((int(520*k),int((y_cmt_bg-46))),tch_name,fill='#595757',font=self.font('丁永康硬笔楷书',35)) 

            draw.text((int(380*k),int((y_logo+10))), '长按二维码 → \n关注视频号 →', fill = '#8E9184',font=self.font('微软雅黑',25))

            
            bg=bg.convert('RGB')
            savename=os.path.join(self.term_pic_dir,std_name+'-'+cmt_date+'-'+term+'阶段小结.jpg')
            bg.save(savename,quality=90,subsampling=0)
            # bg.show()
            print('图片保存完成，保存路径：{}'.format(savename))

        draw_crs()
        # crs_pic()

    def exp_a4_16(self,std_name='韦宇浠',start_date='20200930',end_date='20210121',cmt_date='20210407', \
                        tb_list=[['2020秋','w2'],['2021春','w4']],tch_name='阿晓老师',mode='only16',k=1):
            print('正在处理……')
            info=self.get_std_term_crs(std_name=std_name,tb_list=tb_list,start_date=start_date,end_date=end_date)
            # print(info)
            total_crs=info['total_crs']
            total_crs.dropna(inplace=True)

            std_crs=info['std_crs']
            std_crs.dropna(inplace=True)   
            std_info=info['std_info']

            # print('std_crs',std_crs)
            #如有补课的信息
            if not std_crs[std_crs['课程名称'].str.match('补_\d{8}-.*')].empty:
                missed_crs=std_crs[std_crs['课程名称'].str.match('补_\d{8}-.*')]
                # print('missed_crs',missed_crs)
                #处理该同学的总课表
                for id in missed_crs['index'].tolist():
                    total_crs.drop(total_crs[total_crs['课程名称']==missed_crs[missed_crs['index']==id]['课程名称'].tolist()[0][9:]].index,inplace=True)
                    total_crs.loc[total_crs['上课日期']==missed_crs[missed_crs['index']==id]['上课日期'].tolist()[0],'课程名称']=missed_crs[missed_crs['index']==id]['课程名称'].tolist()[0][11:]

                    total_crs.loc[total_crs['课程名称']==missed_crs[missed_crs['index']==id]['课程名称'].tolist()[0][11:],'上课日期']=missed_crs[missed_crs['index']==id]['课程名称'].tolist()[0][2:11]
                total_crs.sort_values(by=['上课日期'],inplace=True)
                total_crs.drop_duplicates(subset=['上课日期','课程名称'],inplace=True)

                #处理该同学的个人课表              
                std_crs['tmp']=pd.to_datetime(std_crs['课程名称'].apply(lambda x:'-'.join([x[2:6],x[6:8],x[8:10]]) if x.startswith('补') else np.nan))
                std_crs=std_crs[(std_crs['tmp'].isnull()) | (std_crs['tmp']==std_crs['上课日期'])][['上课日期','课程名称']]
                std_crs['课程名称']=std_crs['课程名称'].apply(lambda x:x[11:] if x.startswith('补') else x)

                # print('std_crs',std_crs)

            # print('total_crs',total_crs)



            if mode=='only16':
                # if len(std_crs)<16:
                #     print('{} 在{} - {} 期间总共上了{}节课，未上满16节。未生成报告。'.format(std_name,start_date,end_date,len(std_crs)))
                #     exit(0)
                total_crs=std_crs
            #学期课程总数            
                total_crs_num=16
                crs_lines=2
            else:
            #学期课程总数
                total_crs_num=total_crs.shape[0] 
                # 格子的行数
                if total_crs_num/8==total_crs_num//8: 
                    crs_lines=total_crs_num//8 
                else:
                    crs_lines=total_crs_num//8+1
               

            #所有课程格子的总高度
            ht_crs=int((316+30)*crs_lines)

            #学生基本信息
            std_school=std_info['机构'].values[0]
            std_class=std_info['班级'].values[0]

            #学生学期末评语
            term,weekday=tb_list[-1][0],tb_list[-1][1][1]
            if tb_list[-1][1][0].lower()=='w':
                wd=days_calculate.num_to_ch(weekday)
                xls=os.path.join(self.feedback_dir,'反馈表',term[0:4],term+'-学生课堂学习情况反馈表（周'+wd+'）.xlsx')
            else:
                xls=os.path.join(self.feedback_dir,'反馈表',term[0:4],term+'-学生课堂学习情况反馈表（'+std_name+'）.xlsx')
            comments=WashData.std_feedback(std_name=std_name,xls=xls)['df_term_comment']    
            comments_for_std=comments[comments['学生姓名']==std_name][term+'学期总结-能力-'+str(cmt_date)].values.tolist()[0].replace('#',std_name)
            comments_for_std_psycho=comments[comments['学生姓名']==std_name][term+'学期总结-心理-'+str(cmt_date)].values.tolist()[0].replace('#',std_name)
            # print(std_crs)
            std_crs_num=std_crs.shape[0]
            # if  mode=='only16':
            #     if  total_crs_num==std_crs_num:            
            #         std_crs_num_txt='{0}同学在上一阶段的{1} 节科学机器人课中，完成了全部课程的学习！'.format(std_name,total_crs_num)
            #     else:
            #         std_crs_num_txt='{0}同学在上一阶段的{1} 节科学机器人课中，完成了{2} 节课的学习，请假{3} 节。'.format(std_name,total_crs_num,std_crs_num,total_crs_num-std_crs_num)
            # elif mode=='all':
            #     # if  total_crs_num==std_crs_num:            
            #     std_crs_num_txt='{0}已经完成了{1} 节科学机器人课程的学习！'.format(std_name,total_crs_num)
                # else:
                    # std_crs_num_txt='{0}同学在上一阶段的{1} 节科学机器人课中，完成了{2} 节课的学习，请假{3} 节。'.format(std_name,total_crs_num,std_crs_num,total_crs_num-std_crs_num)
        
            # comments_for_std=std_crs_num_txt+'\n'+comments_for_std

            font_size_cmt=int(50*k)
            prgh_nums=TxtFormat().split_txt_Chn_eng(wid=int(636*k),font_size=font_size_cmt,txt_input=comments_for_std)['para_num']
            #评语标题高度
            ht_cmt_title=int(60*k)
            #评语高度
            ht_prgh=int(font_size_cmt*1.5*(prgh_nums+4))            

            ht_title=int(110*k)
            ht_std=int(100*k)
            ht_bottom=int(210)
            gap_0=int(12*k)
            gap_1=int(13*k)

            y_std_1=int(120*k)
            y_std_2=int((y_std_1+100))
            

            blocks=[ht_title,ht_std,ht_crs,ht_cmt_title,ht_prgh,ht_bottom]

            #计算总的高度
            # ht_total_bg=int((sum(blocks)+gap_0*4+gap_1*(prgh_nums-1)))
            ht_total_bg=3508

            #课程标题坐标
            y_crs_title=int((y_std_2+gap_0+80))
            # y_crs_title=

            #课程格最左上角坐标        
            y_left_up=800

            #评语绿底右下坐标
            y_cmt_bg=y_left_up+ht_crs+gap_0+ht_cmt_title+ht_prgh+30
            # draw.rectangle([(int(36*k),y_left_up+gap_0+ht_cmt_title+10),(int(684*k),y_cmt_bg)],fill='#F3F3E4')

            def draw_bg():
                bg=Image.new('RGBA',(2481,3508),'#7197b4')
                draw=ImageDraw.Draw(bg)
                #白底
                draw.rectangle([(50,54),(2431,3454)],fill='#FFFFFF')
                #评论大框
                draw.rectangle([(166,1720),(2315,3202)],fill='#FFFFFF',outline='#7197b4',width=2)
                #玫瑰图底色
                # draw.rectangle([(1595,1936),(2295,3155)],fill='#f0f8fb')
                # draw.rectangle([(1635,2695),(2255,2855)],fill='#ffffff')
                # draw.rectangle([(1635,2975),(2255,3125)],fill='#ffffff')

                #左边能力底色
                draw.rectangle([(206,1740),(1246,3182)],fill='#f0f8fb')
                draw.rectangle([(226,2537),(1226,3162)],fill='#ffffff')


                #右边心理底色
                draw.rectangle([(1275,1740),(2295,3182)],fill='#f0f8fb')
                draw.rectangle([(1295,2537),(2275,3162)],fill='#ffffff')

                return bg   


            def crs_pic(total_crs=total_crs,std_crs=std_crs,odr=0):
                wid_crs,ht_crs=int(270*k),int(351*k)

                #学生上过的课程和总课程的差集
                sub_set=total_crs.copy()
                sub_set=sub_set.append(std_crs)
                sub_set=sub_set.append(std_crs)
                sub_set.drop_duplicates(subset=['课程名称'],keep=False,inplace=True)

                crs_bg=Image.new('RGB',(wid_crs,ht_crs),'#7197b4')
                crs_name=total_crs.iloc[odr,:]['课程名称']
                crs_date=total_crs.iloc[odr,:]['上课日期']
                crs_pic=Image.open(os.path.join(self.design_dir,crs_name,crs_name[4:]+'.jpg'))
                crs_pic=crs_pic.crop((560,0,1280,720)).resize((int(230*k),int(230*k)))
                crs_bg.paste(crs_pic,(int(20*k),int(20*k)))
                draw=ImageDraw.Draw(crs_bg)
                ft_size=int(30*k)
                x_txt=int(((270*k-len(crs_name[4:])*ft_size)/2+3))
                
                draw.text((x_txt,int(270*k)),crs_name[4:],fill='#ffffff',font=self.font('方正韵动粗黑简',ft_size)) #课程名称
                # draw.text((int(55*k),int(310*k)),str(crs_date)[:11],fill='#ffffff',font=self.font('方正韵动粗黑简',int(24*k))) #课程日期

                #判断是否有不上的课，如有，变灰
                if crs_name not in sub_set['课程名称'].values:
                    draw.text((int(55*k),int(310*k)),str(crs_date)[:11],fill='#ffffff',font=self.font('方正韵动粗黑简',int(24*k))) #课程日期
                else:
                    draw.text((int(55*k),int(310*k)),'---------',fill='#ffffff',font=self.font('方正韵动粗黑简',int(24*k))) #课程日期
                    
                    crs_bg=ImageEnhance.Brightness(crs_bg).enhance(1.4)
                    crs_bg=crs_bg.convert('L')
            
            # crs_bg.show()
                return crs_bg

            def draw_crs():
                print('正在生成图片……',end='')

                #使用并改变母函数的变量
                nonlocal y_left_up
                bg=draw_bg()
                wid_crs,ht_crs=int(270*k),int(351*k)            

                draw=ImageDraw.Draw(bg)
                #分隔线
                # draw.rectangle([(int(36*k),y_crs_title),(int(684*k),y_crs_title+2*k)],fill='#F3F3E4')

                cnt=0
                for i in range(crs_lines):
                    x_begin=int(80*k)              
                    for j in range(8):       
                        if cnt<total_crs_num:  
                            try:
                                crs_block=crs_pic(total_crs=total_crs,std_crs=std_crs,odr=cnt)       
                                bg.paste(crs_block,(x_begin,y_left_up))
                                x_begin=x_begin+wid_crs+int(23*k)   
                                cnt+=1   
                            except Exception as e:
                                pass
                            continue
                    y_left_up=y_left_up+ht_crs+gap_1 
                
                
                #分隔线
                # draw.rectangle([(int(36*k),y_left_up+gap_0),(int(684*k),y_left_up+gap_0+2)],fill='#F3F3E4')
                #评语绿底
                # y_cmt_bg=y_left_up+gap_0+ht_cmt_title+ht_prgh
                # draw.rectangle([(int(36*k),y_left_up+gap_0+ht_cmt_title+10),(int(684*k),y_cmt_bg)],fill='#F3F3E4')

                #二维码，logo
                _pic_logo=Image.open(os.path.join(self.public_dir,'图片类','大智小超新logo.png')).convert('RGBA')
                pic_logo=_pic_logo.resize((int(330*k),int(330*k/2.76)))
                r,g,b,a=pic_logo.split()

                y_logo=int((ht_total_bg-240))

                _pic_qrcode=Image.open(os.path.join(self.public_dir,'图片类','大智小超视频号二维码2.png'))
                pic_qrcode=_pic_qrcode.resize((int(210*k),int(210*k)))

                bg.paste(pic_logo,(int(130*k),int(y_logo+10)),mask=a)
                bg.paste(pic_qrcode,(int(2190*k),int(y_logo-40)))

                #玫瑰图
                rose_and_bar=self.rose_and_bar(std_name=std_name,xls=xls)
                rose_mat=rose_and_bar['res_rose']['chart']
                rose_pic=Plot().mat_to_pil_img(rose_mat)
                rose_pic=rose_pic.resize((640,640*rose_pic.size[1]//rose_pic.size[0]))
                rose_pic=rose_pic.crop((0,0,620,640))
                bg.paste(rose_pic,(421,1890))

                rose_data=rose_and_bar['res_rose']['data']
                
                #柱形图
                # bar_mat=rose_and_bar['res_bar']['chart']
                # bar_pic=pic_transfer.mat_to_pil_img(bar_mat)
                # bar_pic=bar_pic.resize((640,640*rose_pic.size[1]//rose_pic.size[0]))
                # bar_pic=bar_pic.crop((0,0,620,640))
                # bg.paste(bar_pic,(1460,1890))

                #雷达图
                radar_mat=rose_and_bar['res_radar']['chart']
                radar_pic=Plot().mat_to_pil_img(radar_mat)
                radar_pic=radar_pic.resize((640,640*radar_pic.size[1]//radar_pic.size[0]))
                # radar_pic=radar_pic.crop((0,0,620,640))
                bg.paste(radar_pic,(1460,1890))

                #签名(图片)
                tch_sig_txt=[]
                for m,tch_n in enumerate(tch_name):
                    try:
                        pic_tch_sig=Image.open(os.path.join(self.tch_sig_pic_dir,tch_n+'.png'))
                        pic_tch_sig=pic_tch_sig.resize((200,200*pic_tch_sig.size[1]//pic_tch_sig.size[0]))
                        r_sig,g_sig,b_sib,a_sig=pic_tch_sig.split()
                        bg.paste(pic_tch_sig,(950+m*1000,3015),mask=pic_tch_sig)
                    except FileNotFoundError as e:
                        # print(e)
                        tch_sig_txt.append([m,tch_n])
                        # print(tch_n,e)

                # print('签名图片不存在',tch_sig_txt)
                #最高和高低的两个能力名称
                abl_btm=[rose_data[0][0],rose_data[1][0]]
                abl_top=[rose_data[3][0],rose_data[4][0]]

                #文字部分
                #标题
                draw.text((int(1020*k),int(101*k)),'科学机器人课',fill='#7197b4',font=self.font('方正韵动粗黑简',int(78*k))) 
                draw.text((int(890*k),int(220*k)),'课程学习报告',fill='#7197b4',font=self.font('方正韵动粗黑简',int(120*k))) 
                #姓名班级
                ftsz_name=120
                x_name=(2481-ftsz_name*len(std_name))//2
                draw.text((int(x_name),int(421*k)),std_name,fill='#3e3a39',font=self.font('汉仪字酷堂义山楷w',int(ftsz_name*k))) 
                # draw.text((int(400*k*0.9),int(182*k*0.9)),std_school+' '+std_class,fill='#6C8466',font=self.font('汉仪字酷堂义山楷w',int(30*k))) 

                #课程信息

                # print(total_crs)
                prd=total_crs['上课日期'].apply(lambda x:x.strftime('%Y-%m-%d')).tolist()
                mth_start=prd[0].split('-')[0]+'年'+prd[0].split('-')[1]+'月'
                mth_end=prd[-1].split('-')[0]+'年'+prd[-1].split('-')[1]+'月'
                prd_txt='{}-{}'.format(mth_start,mth_end)
                std_crs_finish_number=len(std_crs) if len(std_crs)<=16 else 16
 
                draw.text((930,580),prd_txt,fill='#7197b4',font=self.font('方正韵动粗黑简',int(50*k))) 
                # draw.text((820,730),'在大智小超科学实验室学习',fill='#7197b4',font=self.font('方正韵动粗黑简',int(50*k))) 
                draw.text((700,670),'完成了{}节科学机器人课程的学习'.format(std_crs_finish_number),fill='#7197b4',font=self.font('方正韵动粗黑简',int(72*k))) 
                draw.text((980,1594),'学习情况总结',fill='#7197b4',font=self.font('优设标题',int(90*k))) 
                #能力描述
                # format(abl_top[0],abl_top[1],abl_btm[0],abl_btm[1])
                # draw.text((1790,2625),'你现在最棒的能力',fill='#9f9c9c',font=self.font('优设标题',int(50*k))) 
                # draw.text((1750,2905),'还可以再提升的能力',fill='#9f9c9c',font=self.font('优设标题',int(50*k))) 
                # draw.text((1750,2715),'• '+abl_top[1],fill=rose_data[-1][2],font=self.font('微软雅黑',int(70*k))) 
                # draw.text((1750,2995),'• '+abl_btm[0],fill=rose_data[0][2],font=self.font('微软雅黑',int(70*k)))                

                # TxtFormat.put_txt_img(draw=draw,tt='• '+abl_top[0]+'\n'+'• '+abl_top[1],total_dis=int(800*k*0.95), \
                #                     xy=[1750,2705],dis_line=int(23*k),fill='#7197b4', \
                #                     font_name='微软雅黑',font_size=40,addSPC="yes")
                # TxtFormat.put_txt_img(draw=draw,tt='• '+abl_btm[0]+'\n'+'• '+abl_btm[1],total_dis=int(800*k*0.95), \
                #                     xy=[1750,2995],dis_line=int(23*k),fill='#7197b4', \
                #                     font_name='微软雅黑',font_size=40,addSPC="yes")

                #评语
                #左右标题
                draw.text((518,1764),'搭建能力成长',fill='#7197b4',font=self.font('优设标题',int(80*k)))
                draw.text((1498,1764),'社会性与情感发展',fill='#7197b4',font=self.font('优设标题',int(80*k)))
                draw.text((698,3034),'上课老师：',fill='#7197b4',font=self.font('优设标题',int(50*k)))
                draw.text((1618,3034),'特约心理老师：',fill='#7197b4',font=self.font('优设标题',int(50*k)))


                TxtFormat().put_txt_img(draw=draw,tt=comments_for_std,total_dis=int(980*k*0.95), \
                                    xy=[248,2607],dis_line=int(23*k),fill='#3e3a39', \
                                    font_name='丁永康硬笔楷书',font_size=font_size_cmt,addSPC="yes")

                TxtFormat().put_txt_img(draw=draw,tt=comments_for_std_psycho,total_dis=int(980*k*0.95), \
                                    xy=[1313,2607],dis_line=int(23*k),fill='#3e3a39', \
                                    font_name='丁永康硬笔楷书',font_size=font_size_cmt,addSPC="yes")
                

                #签名(如无图片，则用文字)
                #上课
                # draw.text((950,3035),tch_name[0]+'老师',fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(50*k))) 
                draw.text((950,3100),cmt_date[0:4]+'.'+cmt_date[4:6]+'.'+cmt_date[6:],fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(40*k))) 
                #心理
                # draw.text((1950,3035),tch_name[1]+'老师',fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(50*k))) 
                draw.text((1950,3100),cmt_date[0:4]+'.'+cmt_date[4:6]+'.'+cmt_date[6:],fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(40*k))) 

                if tch_sig_txt:
                    for tch_sig_t in tch_sig_txt:
                        if tch_sig_t[0]==0:
                            draw.text((950,3035),tch_sig_t[1]+'老师',fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(50*k))) 
                        elif tch_sig_t[0]==1:
                            draw.text((1950,3035),tch_sig_t[1]+'老师',fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(50*k))) 
                        else:
                            print('没有该老师签名，请检查。')


                #slogan
                draw.text((600,3300),'让孩子从小玩科学，爱科学，学科学。',fill='#c23030',font=self.font('楷体',int(82*k))) 

                # draw.text((int(380*k),int((y_logo+10))), '长按二维码 → \n关注视频号 →', fill = '#8E9184',font=self.font('微软雅黑',25))

                
                bg=bg.convert('RGB')
                savename=os.path.join(self.term_pic_dir,std_name+'-'+cmt_date+'-'+term+'-课程学习报告.jpg')
                if not os.path.exists(self.term_pic_dir):
                    os.makedirs(self.term_pic_dir)
                # savename=os.path.join('e:/temp/',std_name+'.jpg')
                bg.save(savename,quality=90,subsampling=0)
                # bg.save(savename,quality=95,dpi=(300,300))
                # bg.show()
                print('图片保存完成，保存路径：{}'.format(savename))

            draw_crs()
            # crs_pic()     

    def rose_and_bar(self,std_name,xls):
        # print(xls)
        df=WashData.std_feedback(std_name=std_name,xls=xls)
        df_ability=df['df_ability']
        std_ability=df_ability[df_ability['学生姓名']==std_name].iloc[:,1:]
        dat=std_ability.columns.tolist()
        score=std_ability.iloc[0,:].tolist()
        dat_skill=dat[:5]
        score_skill=score[:5]
        dat_psycho=dat[5:]
        score_psycho=score[5:]

        def rose():        
            # dat=['理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力','协作能力']
            # score=[3,3,4,5,2,2,3,5]
            # score.sort()

            # colors=['#0D9482','#7594CA','#E37933','#E24B29','#F4BA19','#D199C3','#6EA647','#21A4DE']
            colors=['#0D9482','#7594CA','#E37933','#E24B29','#F4BA19']
            df_rose=pd.DataFrame({'dat':dat_skill,'score':score_skill,'colors':colors})

            #排序
            df_rose.sort_values(by=['score'],inplace=True)

            theta = np.linspace(0,2*np.pi,len(dat_skill),endpoint=False)    # 360度等分成n份,endpoint设置是否封闭

            # 设置画布
            fig = plt.figure(figsize=(12,12))
            # 极坐标
            ax = plt.subplot(111,projection = 'polar')
            # 顺时针并设置N方向为0度
            ax.set_theta_direction(-1)
            ax.set_theta_zero_location('N') 

            # 在极坐标中画柱形图
            
            ax.bar(theta,
                    df_rose['score'].values,
                    width = 2*np.pi/5,
                    # color = np.random.random((len(score),3)),
                    # color=['#1E4D58','#245C6A','#296C7C','#2F7B8D','#348B9F','#399AB1','#40A9C2','#51B1C8'],
                    color=df_rose['colors'],
                    # labels=str(country_list), 
                    align = 'edge')

            ## 绘制中心空白
            ax.bar(x=0,    # 柱体的角度坐标
                height=0.8,    # 柱体的高度, 半径坐标
                width=np.pi*2,    # 柱体的宽度
                color='white'
                )
            ''' 
                显示一些简单的中文图例
            '''
            plt.rcParams['font.sans-serif']=['SimHei']  # 黑体
            title=std_name+'同学能力测评'
            # ax.set_title(title,fontdict={'fontsize':20,'color':'#3923a8'})
            #数据标签坐标附加系数
            # y_k=[0.5,0.4,0.3,0.6,1.2,2.4,3.0,1.5]
            y_k=[0.5,0.4,0.9,3.3,2.1]
            angles=[x*np.pi/5 for x in range(1,2*len(dat_skill),2)]
            angles[-1]=angles[-1]-np.pi/16
            # angles[6],angles[7]=angles[6]-np.pi/16,angles[7]-np.pi/16
            for angle,scores,data,color,k in zip(angles,df_rose['score'].values,df_rose['dat'].values,df_rose['colors'].values,y_k):
                ax.text(angle,scores+k,data,color=color,fontsize=28) 
                # ax.text(angle+0.04,scores+0.6,str(scores))

            ax.axis('off')
            # plt.savefig('e:/temp/Nightingale_rose_wcy.jpg',pil_kwargs={'quality':90},dpi=300)
            # plt.show()
            return {'chart':fig,'data':df_rose.values}
        
        def bar():
            fig = plt.figure(figsize=(12,12))
            # 极坐标
            ax = plt.subplot(111)
            plt.rcParams['font.sans-serif']=['SimHei']  # 黑体
            colors=['#D199C3','#6EA647','#21A4DE']
            ax.bar(range(len(dat_psycho)),score_psycho,color=colors)            
            x_major_locator = MultipleLocator(1)
            ax.xaxis.set_major_locator(x_major_locator)
            ax.set_xticks(range(len(dat_psycho)))
            ax.set_xticklabels(dat_psycho,fontsize=24)
            ax.get_xticklabels()[0].set_color("#D199C3")
            ax.get_xticklabels()[1].set_color("#6EA647")
            ax.get_xticklabels()[2].set_color("#21A4DE")

            ax.set_yticks([])
            ax.spines['right'].set_color('none')
            ax.spines['top'].set_color('none')
            ax.spines['left'].set_color('none')
            # plt.xticks([0,1,2],dat_psycho)

            # plt.show()
            return {'chart':fig ,'data':''}

        def radar():
            xxl,rjl,qxl,zyl,zkl=dat_psycho
            # 构造数据
            values = score_psycho
            feature = dat_psycho

            N = len(values)
            # 设置雷达图的角度，用于平分切开一个圆面
            angles = np.linspace(0, 2 * np.pi, N, endpoint=False)


            # 为了使雷达图一圈封闭起来，需要下面的步骤
            values = np.concatenate((values, [values[0]]))
            angles = np.concatenate((angles, [angles[0]]))

            # print(values,angles)

            # 绘图
            fig = plt.figure(figsize=(12,12))
            # 这里一定要设置为极坐标格式
            ax = fig.add_subplot(111, polar=True)
            plt.rcParams['font.sans-serif']=['SimHei']  # 黑体
            # ccl=ax.patch
            
            colors=['#A85727','#54F537','#F58E51','#4438F5','#7976A8']
            # 绘制折线图
            ax.plot(angles, values, '.-', ms=30,linewidth=3,color='#F5EB02')
            # 填充颜色
            ax.fill(angles, values, color='#F5EB02',alpha=0.25)
            # 添加每个特征的标签
            ax.set_thetagrids(angles * 180 / np.pi, '',color='r',fontsize=13)
            # 设置雷达图的范围
            r_distance=10
            ax.set_rlim(0, r_distance)

            ax.grid(color='g', alpha=0.25, lw=3)
            ax.spines['polar'].set_color('grey')
            ax.spines['polar'].set_alpha(0.2)
            ax.spines['polar'].set_linewidth(3)
            # ax.spines['polar'].set_linestyle('-.')

            #项目名称：
            a=[0,0,np.pi/30,-np.pi/50,0,0,0]
            # b=[r_distance*1.1,r_distance*1.1,r_distance*1.2,r_distance*1.4,r_distance*1.5,r_distance*1.2,r_distance*1.1]
            # b=[r_distance*0.8,r_distance*1.1,r_distance*1.4,r_distance*1.4,r_distance*1.1]

            # b=[r_distance*0.8,r_distance*1.1,r_distance*1.4,r_distance*1.4,r_distance*1.1]
            b=[score_psycho[0]*1.05,score_psycho[1]*1.05,score_psycho[2]*1.35,score_psycho[3]*1.45,score_psycho[4]*1.1]

            feature2 = dat_psycho
            for k,i in enumerate(angles):
                try:
                    # print(k,i,feature2[k])
                    ax.text(i+a[k],b[k],feature2[k],fontsize=34,color=colors[k])
                    # ax.text(i+a[k],score_psycho[k],feature2[k],fontsize=34,color=colors[k])
                    
                except:
                    pass

            #分值：
            # c = [1, 0.6, 1.6, 2.3, 1.5, 1,1]
            # print(len(angles))
            # for j,i in enumerate(angles):
            #     try:
            #         r=values[j]-2*i/np.pi
            #         ax.text(i,values[j]+c[j],values[j],color='#218FBD',fontsize=18)
            #     except:
            #         pass

            # 添加标题
            #plt.title('活动前后员工状态表现')

            # 添加网格线
            ax.grid(True,color='grey',alpha=0.05)

            # a=np.arange(0,2*np.pi,0.01)
            # ax.plot(a,10*np.ones_like(a),linewidth=2,color='b')

            ax.set_yticklabels([])
            # plt.savefig(savefilename,transparent=True,bbox_inches='tight')
            # 显示图形
            
            # plt.show()

            return {'chart':fig,'data':''}
        
        res_rose=rose()        
        res_bar=bar()
        res_radar=radar()
        # print(len(score))
        # print(df)
        return {'res_rose':res_rose,'res_bar':res_bar,'res_radar':res_radar}



if __name__=='__main__':
    my=data_summary()
    # res=my.get_std_term_crs(std_name='韦华晋',tb_list=[['2020秋','w6'],['2021春','w6']],start_date='20200901',end_date='20210521')
    # print(res['total_crs'],'\n',res['std_crs'])
    # pic=my.rose_and_bar(std_name='韦宇浠',xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\每周课程反馈\\2021春-学生课堂学习情况反馈表（周四）.xlsx')
    # k=my.get_std_term_crs(std_name='韦万祎',tb_list=[['2021秋','韦万祎']],start_date='20210915',end_date='20220120')
    # print(k)


    # ns=['韦万祎']
    # for n in ns:
    #     my.exp_a4_16(std_name=n,start_date='20210801',end_date='20211110', \
    #                 cmt_date='20210719',tb_list=[['2021秋','w5']], \
    #                 tch_name=['阿晓','杨芳芳'],mode='all',k=1)
        # info=my.get_std_term_crs(std_name=n,tb_list=[['2021秋','w5']],start_date='20210801',end_date='20211110')
        # print(info['total_crs'])
        # print(info['std_crs'])
        # print(info['std_info'])