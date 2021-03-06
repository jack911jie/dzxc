import os
import sys
# print(os.path.dirname(os.path.dirname(os.path.realpath(__file__))))
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'module'))
import composing
import pic_transfer
import WashData
import days_calculate
from readConfig import readConfig
import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image,ImageDraw,ImageFont,ImageEnhance

class data_summary:
    def __init__(self):
        config=readConfig(os.path.join(os.path.dirname(os.path.realpath(__file__)),'configs','TermSummary.config'))
        self.std_dir=config['学生信息文件夹']
        self.design_dir=config['图纸文件夹']
        self.feedback_dir=config['学生课程反馈表文件夹']
        self.public_dir=config['公共素材文件夹']
        self.term_pic_dir='I:\\乐高\\学员上课总结\\学员阶段总结'

        font_cfg=readConfig(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'configs','dzxc_fonts.config'))
        self.font_list=font_cfg['fontList']
        self.font=composing.fonts
        
    def get_std_term_crs(self,std_name='韦宇浠',tb_list=[['2020秋','w2'],['2021春','w4']],start_date='20200901',end_date='20210521'):

        print('    正在读取 {} 的数据……'.format(std_name),end='')
        std_name=std_name.strip()
        std_term_crs=[]
        for tb in tb_list:
            term,weekday=tb[0],tb[1][-1]
            xls_name=os.path.join(self.std_dir,'学生信息表',term+'-学生信息表（周'+days_calculate.num_to_ch(weekday)+'）.xlsx') 
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

        # print(df_std_crs,'\n',df_total_crs)
        # print(df_total_crs)

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
        prgh_nums=composing.split_txt_Chn_eng(wid=int(636*k),font_size=font_size_cmt,txt_input=comments_for_std)[1]
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
            composing.put_txt_img(draw=draw,tt=comments_for_std,total_dis=int(628*k*0.95), \
                                  xy=[int(58),int((y_left_up+gap_0+110))],dis_line=int(20*k),fill='#595757', \
                                  font_name='丁永康硬笔楷书',font_size=font_size_cmt,addSPC="add_2spaces")

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
            total_crs=info['total_crs']
            total_crs.dropna(inplace=True)
            std_crs=info['std_crs']
            std_crs.dropna(inplace=True)   
            std_info=info['std_info']

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
            wd=days_calculate.num_to_ch(weekday)
            xls=os.path.join(self.feedback_dir,term+'-学生课堂学习情况反馈表（周'+wd+'）.xlsx')
            comments=WashData.std_feedback(std_name=std_name,xls=xls)['df_term_comment']    
            comments_for_std=comments[comments['学生姓名']==std_name][term+'学期总结-'+str(cmt_date)].values.tolist()[0].replace('#',std_name)
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
            prgh_nums=composing.split_txt_Chn_eng(wid=int(636*k),font_size=font_size_cmt,txt_input=comments_for_std)[1]
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
            y_left_up=900

            #评语绿底右下坐标
            y_cmt_bg=y_left_up+ht_crs+gap_0+ht_cmt_title+ht_prgh+30
            # draw.rectangle([(int(36*k),y_left_up+gap_0+ht_cmt_title+10),(int(684*k),y_cmt_bg)],fill='#F3F3E4')

            def draw_bg():
                bg=Image.new('RGBA',(2481,3508),'#7197b4')
                draw=ImageDraw.Draw(bg)
                #白底
                draw.rectangle([(50,54),(2431,3454)],fill='#FFFFFF')
                #评论大框
                draw.rectangle([(166,1916),(2315,3182)],fill='#FFFFFF',outline='#7197b4',width=2)
                #玫瑰图底色
                draw.rectangle([(1595,1936),(2295,3155)],fill='#f0f8fb')
                draw.rectangle([(1635,2695),(2255,2855)],fill='#ffffff')
                draw.rectangle([(1635,2975),(2255,3125)],fill='#ffffff')

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
                rose=self.rose(std_name=std_name,xls=xls)
                rose_mat=rose['chart']
                rose_pic=pic_transfer.mat_to_pil_img(rose_mat)
                rose_pic=rose_pic.resize((640,640*rose_pic.size[1]//rose_pic.size[0]))
                rose_pic=rose_pic.crop((0,0,620,640))
                bg.paste(rose_pic,(1635,1966))

                rose_data=rose['data']
                
                #最高和高低的两个能力名称
                abl_btm=[rose_data[0][0],rose_data[1][0]]
                abl_top=[rose_data[6][0],rose_data[7][0]]

                #文字部分
                #标题
                draw.text((int(1020*k),int(101*k)),'科学机器人课',fill='#7197b4',font=self.font('方正韵动粗黑简',int(78*k))) 
                draw.text((int(890*k),int(220*k)),'课程学习报告',fill='#7197b4',font=self.font('方正韵动粗黑简',int(120*k))) 
                #姓名班级
                ftsz_name=120
                x_name=(2481-ftsz_name*len(std_name))//2
                draw.text((int(x_name),int(481*k)),std_name,fill='#3e3a39',font=self.font('汉仪字酷堂义山楷w',int(ftsz_name*k))) 
                # draw.text((int(400*k*0.9),int(182*k*0.9)),std_school+' '+std_class,fill='#6C8466',font=self.font('汉仪字酷堂义山楷w',int(30*k))) 

                #课程信息

                # print(total_crs)
                prd=total_crs['上课日期'].apply(lambda x:x.strftime('%Y-%m-%d')).tolist()
                mth_start=prd[0].split('-')[0]+'年'+prd[0].split('-')[1]+'月'
                mth_end=prd[-1].split('-')[0]+'年'+prd[-1].split('-')[1]+'月'
                prd_txt='{}-{}'.format(mth_start,mth_end)
                std_crs_finish_number=len(std_crs) if len(std_crs)<=16 else 16
 
                draw.text((930,680),prd_txt,fill='#7197b4',font=self.font('方正韵动粗黑简',int(50*k))) 
                # draw.text((820,730),'在大智小超科学实验室学习',fill='#7197b4',font=self.font('方正韵动粗黑简',int(50*k))) 
                draw.text((700,770),'完成了{}节科学机器人课程的学习'.format(std_crs_finish_number),fill='#7197b4',font=self.font('方正韵动粗黑简',int(72*k))) 
                draw.text((980,1774),'学习情况总结',fill='#7197b4',font=self.font('优设标题',int(90*k))) 
                #能力描述
                # format(abl_top[0],abl_top[1],abl_btm[0],abl_btm[1])
                draw.text((1790,2625),'你现在最棒的能力',fill='#9f9c9c',font=self.font('优设标题',int(50*k))) 
                draw.text((1750,2905),'还可以再提升的能力',fill='#9f9c9c',font=self.font('优设标题',int(50*k))) 
                draw.text((1750,2715),'• '+abl_top[1],fill=rose_data[-1][2],font=self.font('微软雅黑',int(70*k))) 
                draw.text((1750,2995),'• '+abl_btm[0],fill=rose_data[0][2],font=self.font('微软雅黑',int(70*k)))                

                # composing.put_txt_img(draw=draw,tt='• '+abl_top[0]+'\n'+'• '+abl_top[1],total_dis=int(800*k*0.95), \
                #                     xy=[1750,2705],dis_line=int(23*k),fill='#7197b4', \
                #                     font_name='微软雅黑',font_size=40,addSPC="add_2spaces")
                # composing.put_txt_img(draw=draw,tt='• '+abl_btm[0]+'\n'+'• '+abl_btm[1],total_dis=int(800*k*0.95), \
                #                     xy=[1750,2995],dis_line=int(23*k),fill='#7197b4', \
                #                     font_name='微软雅黑',font_size=40,addSPC="add_2spaces")

                #评语
                composing.put_txt_img(draw=draw,tt=comments_for_std,total_dis=int(1343*k*0.95), \
                                    xy=[218,1955],dis_line=int(23*k),fill='#3e3a39', \
                                    font_name='丁永康硬笔楷书',font_size=font_size_cmt,addSPC="add_2spaces")

                #签名
                draw.text((1200,3025),tch_name,fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(70*k))) 
                draw.text((1250,3120),cmt_date[0:4]+'.'+cmt_date[4:6]+'.'+cmt_date[6:],fill='#3e3a39',font=self.font('丁永康硬笔楷书',int(50*k))) 

                #slogan
                draw.text((600,3300),'让孩子从小玩科学，爱科学，学科学。',fill='#c23030',font=self.font('楷体',int(82*k))) 

                # draw.text((int(380*k),int((y_logo+10))), '长按二维码 → \n关注视频号 →', fill = '#8E9184',font=self.font('微软雅黑',25))

                
                bg=bg.convert('RGB')
                savename=os.path.join(self.term_pic_dir,std_name+'-'+cmt_date+'-'+term+'-课程学习报告.jpg')
                # savename=os.path.join('e:/temp/',std_name+'.jpg')
                # bg.save(savename,quality=90,subsampling=0)
                bg.save(savename,quality=95,dpi=(300,300))
                # bg.show()
                print('图片保存完成，保存路径：{}'.format(savename))

            draw_crs()
            # crs_pic()     

    def rose(self,std_name,xls):
        # print(xls)
        df=WashData.std_feedback(std_name=std_name,xls=xls)
        df_ability=df['df_ability']
        std_ability=df_ability[df_ability['学生姓名']==std_name].iloc[:,1:]
        dat=std_ability.columns.tolist()
        score=std_ability.iloc[0,:].tolist()

        # dat=['理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力','协作能力']
        # score=[3,3,4,5,2,2,3,5]
        # score.sort()

        colors=['#0D9482','#7594CA','#E37933','#E24B29','#F4BA19','#D199C3','#6EA647','#21A4DE']
        df=pd.DataFrame({'dat':dat,'score':score,'colors':colors})

        #排序
        df.sort_values(by=['score'],inplace=True)

        theta = np.linspace(0,2*np.pi,len(dat),endpoint=False)    # 360度等分成n份,endpoint设置是否封闭

        # 设置画布
        fig = plt.figure(figsize=(12,12))
        # 极坐标
        ax = plt.subplot(111,projection = 'polar')
        # 顺时针并设置N方向为0度
        ax.set_theta_direction(-1)
        ax.set_theta_zero_location('N') 

        # 在极坐标中画柱形图
        
        ax.bar(theta,
                df['score'].values,
                width = 0.75,
                # color = np.random.random((len(score),3)),
                # color=['#1E4D58','#245C6A','#296C7C','#2F7B8D','#348B9F','#399AB1','#40A9C2','#51B1C8'],
                color=df['colors'],
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
        y_k=[0.5,0.4,0.3,0.6,1.2,2.4,3.0,1.5]
        angles=[x*np.pi/8 for x in range(1,2*len(dat),2)]
        angles[6],angles[7]=angles[6]-np.pi/16,angles[7]-np.pi/16
        for angle,scores,data,color,k in zip(angles,df['score'].values,df['dat'].values,df['colors'].values,y_k):
            ax.text(angle,scores+k,data,color=color,fontsize=28) 
            # ax.text(angle+0.04,scores+0.6,str(scores))

        ax.axis('off')
        # plt.savefig('e:/temp/Nightingale_rose_wcy.jpg',pil_kwargs={'quality':90},dpi=300)
        # plt.show()
        # print(len(score))
        return {'chart':fig,'data':df.values}



if __name__=='__main__':
    my=data_summary()
    # res=my.get_std_term_crs(std_name='韦华晋',tb_list=[['2020秋','w6'],['2021春','w6']],start_date='20200901',end_date='20210521')
    # print(res['total_crs'],'\n',res['std_crs'])
    # my.rose(std_name='韦宇浠',weekday=2)
    ns=['韦华晋','黄建乐']
    for n in ns:
        my.exp_a4_16(std_name=n,start_date='20200801',end_date='20210510', \
                    cmt_date='20210410',tb_list=[['2020秋','w6'],['2021春','w6']], \
                    tch_name='阿晓老师',mode='only16',k=1)
