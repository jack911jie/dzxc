import os
import sys
sys.path.append('i:/py/dzxc/module')
import composing
import WashData
import days_calculate
from readConfig import readConfig
import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image,ImageDraw,ImageFont,ImageEnhance

class data_summary:
    def __init__(self):
        config=readConfig(os.path.join(os.path.dirname(os.path.realpath(__file__)),'configs','term_summary_config.dazhi'))
        self.std_dir=config['学生信息文件夹']
        self.design_dir=config['图纸文件夹']
        self.feedback_xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\5-超智幼儿园\\每周课程反馈\\学员课堂学习情况反馈表.xlsx'

        font_cfg=readConfig(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'configs','dzxc_fonts.config'))
        self.font_list=font_cfg['fontList']
        self.font=composing.fonts
        
    def draw_std_term_crs(self,std_name='韦宇浠',start_date='20200930',end_date='20210121',weekday='2'):
        std_name=std_name.strip()
        xls_name=os.path.join(self.std_dir,'2020乐高课程签到表（周'+days_calculate.num_to_ch(weekday)+'）.xlsx')
        std_term_crs=WashData.std_term_crs(std_name=std_name,start_date=start_date,end_date=end_date,xls=xls_name)
        return std_term_crs

    def exp_poster(self,std_name='韦宇浠',start_date='20200930',end_date='20210121',weekday='2',term='2020秋',tch_name='阿晓老师'):
        info=self.draw_std_term_crs(std_name=std_name,start_date=start_date,end_date=end_date,weekday=weekday)
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
        ht_crs=(170+14)*crs_lines

        #学生基本信息
        std_school=std_info['机构'].values[0]
        std_class=std_info['班级'].values[0]

        comments=WashData.std_feedback(std_name=std_name,xls=self.feedback_xls,weekday=2)['df_term_comment']
        #学生学期末评语
        comments_for_std=comments[comments['姓名']==std_name][term+'学期总结'].values.tolist()[0].replace('#',std_name)
        std_crs_num=std_crs.shape[0]
        if  total_crs_num==std_crs_num:            
            std_crs_num_txt='{0}同学在上一阶段的{1}节科学机器人课中，完成了全部课程的学习！'.format(std_name,total_crs_num)
        else:
             std_crs_num_txt='{0}同学在上一阶段的{1}节科学机器人课中，完成了{2}节课的学习，请假{3}节。'.format(std_name,total_crs_num,std_crs_num,total_crs_num-std_crs_num)
        comments_for_std=std_crs_num_txt+'\n'+comments_for_std
        font_size_cmt=30
        prgh_nums=composing.split_txt_Chn_eng(wid=636,font_size=font_size_cmt,txt_input=comments_for_std)[1]
        #评语标题高度
        ht_cmt_title=60
        #评语高度
        ht_prgh=font_size_cmt*1.5*(prgh_nums+4)

        

        ht_title=110
        ht_std=100
        ht_bottom=210
        gap_0=12
        gap_1=13

        y_std_1=120
        y_std_2=y_std_1+100
        

        blocks=[ht_title,ht_std,ht_crs,ht_cmt_title,ht_prgh,ht_bottom]
        #计算总的高度
        ht_total_bg=int(sum(blocks)+gap_0*4+gap_1*(prgh_nums-1))

        #课程标题坐标
        y_crs_title=y_std_2+gap_0+80

        #课程格最左上角坐标        
        y_left_up=y_crs_title+gap_0

        def draw_bg():
            bg=Image.new('RGBA',(720,ht_total_bg),'#F3F9E4')
            draw=ImageDraw.Draw(bg)
            #白底
            draw.rectangle([(13,110),(707,ht_total_bg-25)],fill='#FFFFFF')
            #名字绿底
            draw.rectangle([(36,y_std_1),(250,y_std_2)],fill='#F3F9E4')
            draw.rectangle([(36,y_std_2),(684,y_std_2+2)],fill='#F3F9E4')
            
            
            return bg

        def crs_pic(total_crs=total_crs,std_crs=std_crs,odr=0):
            wid_crs,ht_crs=150,170

            #学生上过的课程和总课程的差集
            sub_set=total_crs.copy()
            sub_set=sub_set.append(std_crs)
            sub_set=sub_set.append(std_crs)
            sub_set.drop_duplicates(subset=['课程名称'],keep=False,inplace=True)

            crs_bg=Image.new('RGB',(wid_crs,ht_crs),'#F3F3E4')
            crs_name=total_crs.iloc[odr,:]['课程名称']
            crs_date=total_crs.iloc[odr,:]['上课日期']
            crs_pic=Image.open(os.path.join(self.design_dir,crs_name,crs_name[4:]+'.jpg'))
            crs_pic=crs_pic.crop((560,0,1280,720)).resize((110,110))
            crs_bg.paste(crs_pic,(20,6))
            draw=ImageDraw.Draw(crs_bg)
            ft_size=19
            x_txt=int((150-len(crs_name[4:])*ft_size)/2+3)
            
            draw.text((x_txt,122),crs_name[4:],fill='#8fc31f',font=self.font('方正韵动粗黑简',ft_size)) #课程名称

            #判断是否有不上的课，如有，变灰
            if crs_name not in sub_set['课程名称'].values:
                draw.text((35,148),str(crs_date)[:11],fill='#8fc31f',font=self.font('方正韵动粗黑简',12)) #课程日期
            else:
                draw.text((35,148),'---------',fill='#8fc31f',font=self.font('方正韵动粗黑简',12)) #课程日期
                
                crs_bg=ImageEnhance.Brightness(crs_bg).enhance(1.4)
                crs_bg=crs_bg.convert('L')
            
            # crs_bg.show()
            return crs_bg

        def draw_crs():
            #使用并改变母函数的变量
            nonlocal y_left_up
            bg=draw_bg()
            wid_crs,ht_crs=150,170            

            draw=ImageDraw.Draw(bg)
            #分隔线

            draw.rectangle([(36,y_crs_title),(684,y_crs_title+2)],fill='#F3F3E4')

            cnt=0
            for i in range(crs_lines):
                x_begin=40                
                for j in range(4):       
                    if cnt<total_crs_num:      
                        crs_block=crs_pic(total_crs=total_crs,std_crs=std_crs,odr=cnt)       
                        bg.paste(crs_block,(x_begin,y_left_up))
                        x_begin=x_begin+wid_crs+13   
                        cnt+=1         
                y_left_up=y_left_up+ht_crs+gap_1 
            
            
            #分隔线
            draw.rectangle([(36,y_left_up+gap_0),(684,y_left_up+gap_0+2)],fill='#F3F3E4')
           #评语绿底
            y_cmt_bg=y_left_up+gap_0+ht_cmt_title+ht_prgh+10
            draw.rectangle([(36,y_left_up+gap_0+ht_cmt_title+10),(684,y_cmt_bg)],fill='#F3F3E4')

            #二维码，logo
            _pic_logo=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超新logo.png').convert('RGBA')
            pic_logo=_pic_logo.resize((210,int(210/2.76)))
            r,g,b,a=pic_logo.split()

            y_logo=ht_total_bg-130

            _pic_qrcode=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超视频号二维码2.png')
            pic_qrcode=_pic_qrcode.resize((100,100))

            bg.paste(pic_logo,(40,int(y_logo+10)),mask=a)
            bg.paste(pic_qrcode,(580,int(y_logo)))

            #文字部分
            #标题
            draw.text((130,5),'科学机器人课',fill='#8fc31f',font=self.font('汉仪超级战甲',78)) 
            #姓名班级
            draw.text((60,142),std_name,fill='#F8B62D',font=self.font('汉仪字酷堂义山楷w',60)) 
            draw.text((400,182),std_school+' '+std_class,fill='#6C8466',font=self.font('汉仪字酷堂义山楷w',30)) 
            draw.text((45,y_crs_title-52),'点亮的课程',fill='#8fc31f',font=self.font('鸿蒙印品',35)) 
            draw.text((45,y_left_up+gap_0+20),'学期评语',fill='#8fc31f',font=self.font('鸿蒙印品',35)) 

            #评语
            composing.put_txt_img(draw=draw,tt=comments_for_std,total_dis=628, \
                                  xy=[45,y_left_up+gap_0+110],dis_line=20,fill='#595757', \
                                  font_name='丁永康硬笔楷书',font_size=font_size_cmt,addSPC="add_2spaces")

            draw.text((520,y_cmt_bg-46),tch_name,fill='#595757',font=self.font('丁永康硬笔楷书',35)) 

            draw.text((380,y_logo+10), '长按二维码 → \n关注视频号 →', fill = '#8E9184',font=self.font('微软雅黑',25))

            bg.show()
            # bg=bg.convert('RGB')
            # bg.save('e:/temp/kkkk2.jpg',quality=90)

        draw_crs()
        # crs_pic()
        



    def rose(self,std_name='韦宇浠',weekday=2):
        df=WashData.std_feedback(os.path.join(self.std_dir,'每周课程反馈','学员课堂学习情况反馈表.xlsx'),weekday=weekday)
        df_ability=df['df_ability']
        std_ability=df_ability[df_ability['姓名']==std_name].iloc[:,1:]
        dat=std_ability.columns.tolist()
        score=std_ability.iloc[0,:].tolist()

        # dat=['理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力','协作能力']
        # score=[3,3,4,5,2,2,3,5]
        score.sort()
        theta = np.linspace(0,2*np.pi,len(dat),endpoint=False)    # 360度等分成n份

        # 设置画布
        fig = plt.figure(figsize=(12,10))
        # 极坐标
        ax = plt.subplot(111,projection = 'polar')
        # 顺时针并设置N方向为0度
        ax.set_theta_direction(-1)
        ax.set_theta_zero_location('N') 

        # 在极坐标中画柱形图
        ax.bar(theta,
                score,
                width = 0.75,
                color = np.random.random((len(score),3)),
                # labels=str(country_list), 
                align = 'edge')
        ''' 
            显示一些简单的中文图例
        '''
        plt.rcParams['font.sans-serif']=['SimHei']  # 黑体
        title=std_name+'同学能力测评'
        ax.set_title(title,fontdict={'fontsize':20,'color':'#3923a8'})
        for angle,scores,data in zip(theta,score,dat):
            ax.text(angle+0.1,scores+0.2,data) 
            # ax.text(angle+0.04,scores+0.6,str(scores))

        ax.axis('off')
        # plt.savefig('Nightingale_rose.png')
        plt.show()


        # print(len(score))


if __name__=='__main__':
    my=data_summary()
    # my.rose(std_name='陶盛挺',weekday=2)
    my.exp_poster(std_name='韦宇浠',start_date='20200922',end_date='20201201',weekday='2',term='2020秋',tch_name='阿晓老师')