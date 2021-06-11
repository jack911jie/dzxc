import sys
sys.path.append('i:/py/dzxc/module')
import days_calculate
import WashData
import composing
import readConfig
import os
import numpy as np
import json
import time
import re
import random
import logging
import pandas as pd
from PIL import Image,ImageDraw,ImageFont
from tqdm import tqdm
import exifread
import iptcinfo3

logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(funcName)s-%(lineno)d - %(message)s')
logger = logging.getLogger(__name__)

class poster:
    def __init__(self,weekday=2,term='2021春',place_input='001-超智幼儿园'):
        
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'LegoPoster.config'),'r',encoding='utf-8') as f:
            lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=json.loads(_line)

        self.bg=config['导出文件夹']
        self.font_dir=config['字体文件夹']
        self.default_font=config['默认字体']         
        self.crsList=config['课程信息表']
        self.weekday=weekday
        self.term=term
        self.place_input=place_input
        wd=days_calculate.num_to_ch(str(weekday))
        # self.eachStd=config['个别学员评语表']
        # 个别学员评语表":"E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\学员课堂学习情况反馈表.xlsx",
        cmt_table_dir=config['老师评语表文件夹']
        self.eachStd=os.path.join(cmt_table_dir,place_input,'每周课程反馈',term+'-学生课堂学习情况反馈表（周'+wd+'）.xlsx')
        
        self.picTitleDir=config['课程标题照片文件夹']
        self.picStdDir=config['学生照片文件夹']
        self.ConsDir=config['乐高图纸文件夹']
        self.std_sig_dir=config['学生签到表文件夹']

        # if weekday==2:
        #     self.crsStudent=config['学员签到表w2']
        # elif weekday==6:
        #     self.crsStudent=config['学员签到表w6']
        
        self.crsStudent=os.path.join(self.std_sig_dir,place_input,'学生信息表',term+'-学生信息表（周'+wd+'）.xlsx')

        self.PraiseTxt=['#同学在课堂上的表现非常棒！下次课加油！','下节课继续加油哦！','这节课很有收获，期待你下节课能有更大进步！','#同学的课堂表现非常棒，老师给你点个赞！'] 
        self.picWid=425 #默认照片
        self.comment_dis_line=25 #老师评论的行间距
        self.comment_font_size=36 #老师评论的字体大小
        
    def pic_resize(self,pic,wid):
        w,h=pic.size
        r=h/w        
        if r!=0.75:            
            pic_cut=pic.crop((int(w-h/0.75),0,w,h))
            pic_resized=pic_cut.resize((wid,int(wid*0.75)),Image.ANTIALIAS)            
        else:
            pic_resized=pic.resize((wid,int(wid*r)),Image.ANTIALIAS)
            
        return pic_resized        

    def split_txt(self,dis,font_size,txt_input,Indent='no'):    
        txts=txt_input.splitlines()        
        if Indent=='yes':
            for i,t in enumerate(txts):
                txts[i]=chr(12288)+chr(12288)+t
            
        logging.info(txts)
        spt=0
        # logging.info()
        for t in txts:        
            total_len=(self.char_len(t)+len(txts)*6)*font_size

            if total_len%dis==0:
                _spt=total_len/dis        
            else:
                _spt=total_len//dis+1
            spt=spt+_spt
            
        logging.info(spt)

        zi_per_line=int(dis//font_size)
        logging.info(''.join(['每行字数：',str(zi_per_line)]))
        txt_L=[]

        for txt in txts:
            t_n=[]
            for i in range(0,int(spt)):
                if i==0:
                    if txt[0:zi_per_line]:
                        t_n.append(txt[0:zi_per_line])
                else:
                    if txt[i*zi_per_line:(i+1)*zi_per_line]:
                        t_n.append(txt[i*zi_per_line:(i+1)*zi_per_line])
            txt_L.append(t_n)
                        
            # print(len(txt[i*zi_per_line:(i+1)*zi_per_line]))
        logging.info(txt_L)
        return txt_L
    
    def char_len(self,txt):
        len_s=len(txt)
        len_u=len(txt.encode('utf-8'))
        ziShu_z=(len_u-len_s)/2
        ziShu_e=len_s-ziShu_z
        total=ziShu_z+ziShu_e*0.5    
        return total

    def put_txt_img(self,img,t,total_dis,xy,dis_line,fill,font_name,font_size,addSPC='None'):
        
        fontInput=self.fonts(font_name,font_size)            
        if addSPC=='add_2spaces': 
            ind='yes'
        else:
            ind='no'
            
        # txt=self.split_txt(total_dis,font_size,t,Indent='no')
        txt,p_num=composing.split_txt_Chn_eng(total_dis,font_size,t,Indent=ind)

        # font_sig = self.fonts('丁永康硬笔楷书',40)
        draw=ImageDraw.Draw(img)   
        # logging.info(txt)
        n=0
        for t in txt:              
            m=0
            for tt in t:                  
                x,y=xy[0],xy[1]+(font_size+dis_line)*n
                if addSPC=='add_2spaces':   #首行缩进
                    if m==0:    
                        # tt='  '+tt #首先前面加上两个空格
                        logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                        logging.info(tt)
                        draw.text((x-font_size*0.2,y), tt, fill = fill,font=fontInput) 
                    else:                       
                        logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                        logging.info(tt)
                        draw.text((x,y), tt, fill = fill,font=fontInput)  
                else:
                    logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                    logging.info(tt)
                    draw.text((x,y), tt, fill = fill,font=fontInput)  
 
                m+=1
                n+=1

    def sstr(self,x):
        b=[str(i) for i in x]     
        return b
    
    def fonts(self,font_name,font_size):
        fontList={
            '腾祥金砖黑简':'c:\\windows\\fonts\\腾祥金砖黑简.TTF',
            '汉仪糯米团':'j:\\fonts\\HYNuoMiTuanW.ttf',
            '丁永康硬笔楷书':'j:\\fonts\\2012DingYongKangYingBiKaiShuXinBan-2.ttf',
            '微软雅黑':'i:\\py\\msyh.ttf',
            '鸿蒙印品':'j:\\fonts\\hongMengHei.ttf',
            '优设标题':'j:\\fonts\\yousheTitleHei.ttf',
            '汉仪超级战甲':'j:\\fonts\\HYChaoJiZhanJiaW-2.ttf',
            '汉仪心海行楷w':'j:\\fonts\\HYXinHaiXingKaiW.ttf',
            '华康海报体W12(p)':'j:\\fonts\\HuaKangHaiBaoTiW12-P-1.ttf',
            '汉仪锐智w':'j:\\fonts\\HYRuiZhiW.ttf',
            '杨任东竹石体':'j:\\fonts\\yangrendongzhushi-Regular.ttf',
            '于洪亮硬笔行楷手写体':'j:\\fonts\\YuHongLiangYingBiXingKaiShouXieTiZhengShiBan-1.ttf',
            '汉仪林峰体w':'j:\\fonts\\HYLinFengTiW.ttf',
            '汉仪常江小楷':'j:\\fonts\\HYChangJiangXiaoKaiW.ttf',
            '汉仪字酷堂经解楷体w':'j:\\fonts\\HYZiKuTangJingJieKaiTiW.ttf'
        }

        # ImageFont.truetype('j:\\fonts\\2012DingYongKangYingBiKaiShuXinBan-2.ttf',font_size)


        return ImageFont.truetype(fontList[font_name],font_size)
    
    def PosterDraw(self,crs_nameInput,dateInput,TeacherSig='阿晓老师'):
        crs_name=crs_nameInput[4:]
        crs_code=crs_nameInput[0:4]
        def basic_para():
            print('正在初始化参数……',end='')
            picWid=self.picWid
            r=0.75
            picX1,picX3=10,10
            picX2,picX4=picX1+picWid+20,picX1+picWid+20
            picY1,picY2=720,720
            picY3,picY4=picY1+picWid*r+20,picY1+picWid*r+20

            pic0=pic_xy(picWid,int(picWid*r),440,290) #标题图
            pic1=pic_xy(picWid,int(picWid*r),picX1,picY1) #照片一
            pic2=pic_xy(picWid,int(picWid*r),picX2,picY2) #照片二
            pic3=pic_xy(picWid,int(picWid*r),picX3,picY3) #照片三
            pic4=pic_xy(picWid,int(picWid*r),picX4,picY4) #照片四
            print('完成')
            
            return[picWid,picX1,picX2,picX3,picX4,picY1,picY2,picY3,picY4,pic0,pic1,pic2,pic3,pic4]   

        def sortPics(files):
            newfiles={}
            for fn in files:
                if fn[-3:].lower()=='jpg':
                    with open(fn,'rb') as fd:
                        tag=exifread.process_file(fd)
                        t=str(tag['EXIF DateTimeOriginal']).replace(':','-',2)
                        if t in newfiles.keys():
                            t=t[0:-2]+str(int(t[-2:])+1).zfill(2)
                            newfiles[t]=fn
                        else:
                            newfiles[t]=fn
                        # print('exifInfo:',t,fn)
            newList=[]
            for i in sorted(newfiles):
                newList.append(newfiles[i])
            return newList
        
        def read_excel():
            print('正在读取学员和课程信息……',end='')
            infos=WashData.comments_after_class(cmt_date=str(dateInput),crs_name_input=crs_nameInput,weekday=self.weekday, \
                                                crs_list=self.crsList,std_info=self.crsStudent, \
                                                tch_cmt=self.eachStd)
            print('完成')
            return infos         

        def read_scores(place_input,term,weekday):
            print('正在读取积分信息……',end='')
            wd=days_calculate.num_to_ch(str(weekday))
            xls_this=os.path.join(self.std_sig_dir,self.place_input,'学生信息表',term+'-学生信息表（周'+wd+'）.xlsx')  
            df_this_crs_score=WashData.std_score_this_crs(xls_this)

            place_dir=os.path.join(self.std_sig_dir,self.place_input)
            df_all_scores=WashData.std_all_scores(place_dir)

            print('完成')
            return {'this_score':df_this_crs_score,'all_scores':df_all_scores}
        
        def pic_xy(picWid,picHig,x0,y0):
            #白色矩形
            recX0=x0
            recY0=y0
            recX1=x0+picWid+10
            recY1=y0+picHig+10

            #照片
            picX0=recX0+5
            picY0=recY0+5
            picX1=picX0+picWid
            picY1=picY0+picHig

            return [[int(recX0),int(recY0),int(recX1),int(recY1)],[int(picX0),int(picY0),int(picX1),int(picY1)]]        
        
        def color_list(list_num='202002'):
            colors_202002={
                'title_bg':'#FF9E11',  #乐高机器人课
                'name':'#FFFFFF', # 姓名 年龄
                'crs':'#00b578', #课程
                'pics':'#f7f7f7' ,#照片大框
                'pics_box':'#ffffff',#照片相框
                'pic01':'#ffffff',
                'pic02':'#ffffff',
                'pic03':'#ffffff',
                'pic04':'#ffffff',
                'comments':'#00bde3', #老师评语
                'logo':'#ffffff', #logo
                't_title_ke1':'#ffffff',
                't_title_xue':'#ffffff',
                't_title_ji':'#ffffff',
                't_title_qi':'#ffffff',
                't_title_ren':'#ffffff',
                't_title_ke4':'#ffffff',
                'scores':'#f7f7f7', #积分底色
                'score_left':'#ffffff',
                'score_right':'#ffffff',
                't_name':'#00b578',#姓名
                't_kdgtn':'#00b578',#幼儿园
                't_class':'#00b578',#班级
                't_crs':'#FFE340',#课程名称
                't_diff':'#ffffff',#难度
                't_tool':'#ffffff',#教具
                't_knlg':'#ffffff',#知识点
                't_date':'#ffffff',#日期
                't_scores_left':'#8b988e',#积分文字
                't_scores_right':'#8b988e',#积分文字
                't_scores_title':'#8b988e',#积分标题
                't_tch_cmt_title':'#795022',#老师评语标题
                't_tch_cmt':'#ffffff',#老师评语
                't_tch_sig':'#ffffff',#签名
                't_bottom':'#656564' #二维码文字
            }
            colors_202101={
                'title_bg':'#f7f3c3',  #乐高机器人课
                'name':'#FFFFFF', # 姓名 年龄
                'crs':'#f1f9ee', #课程
                'pics':'#f7f7f7' ,#照片大框
                'pics_box':'#ffffff',#照片相框
                'pic01':'#ffffff',
                'pic02':'#ffffff',
                'pic03':'#ffffff',
                'pic04':'#ffffff',
                'scores':'#f7f7f7', #积分底色
                'score_left':'#ffffff',
                'score_right':'#ffffff',
                'comments':'#f9f9f9', #老师评语
                'logo':'#ffffff', #logo
                't_title_ke1':'#009ce6',
                't_title_xue':'#33b171',
                't_title_ji':'#ef9c10',
                't_title_qi':'#8bbe19',
                't_title_ren':'#dc9ee7',
                't_title_ke4':'#b18046',
                't_name':'#99633f',#姓名
                't_kdgtn':'#99633f',#幼儿园
                't_class':'#99633f',#班级
                't_crs':'#37a751',#课程名称
                't_diff':'#8b988e',#难度
                't_tool':'#8b988e',#教具
                't_knlg':'#8b988e',#知识点
                't_date':'#8b988e',#日期
                't_scores_left':'#cc8e28',#积分文字
                't_scores_right':'#c8bb9b',#积分文字
                't_scores_title':'#b4b4b5',#积分标题
                't_tch_cmt_title':'#b4b4b5',#老师评语标题
                't_tch_cmt':'#795022',#老师评语
                't_tch_sig':'#795022',#签名
                't_bottom':'#656564' #二维码文字
            }

            out=eval('colors_'+list_num)

            return out
        
        def basic_bg(num):
            color=color_list('202101')

            # y1:第一块左上角纵坐标，y1_2:第一块右下角纵坐标，以此类推
            # y及s后面的数字代表：
            # 1——标题
            # 2——姓名、机构
            # 3——课程、知识点
            # 4——照片（2张或4张）
            # 5——课堂情况（老师评论），可变，随评论多少而改变。
            # score——积分
            # 6——logo、二维码

            s1=100
            s2=140
            s3=450
            s4=700
            self.s4=s4
            
            #评论的高度
            s5=(self.comment_font_size+self.comment_dis_line)*self.para_num+160
            self.s5=s5

            #积分的高度
            s_score=380

            #logo的高度
            s6=200
            self.s6=s6
            sprt=10   

            if num>3:   
                total_len=s1+s2+s3+s4+s_score+s5+s6+sprt*6
            else:
                total_len=s1+s2+s3+s4+s_score+s5+s6+sprt*6-s4//2            

            #标题坐标
            y1=0
            y1_2=y1+s1
            #姓名、机构坐标
            y2=y1_2+sprt        
            y2_2=y2+s2
            #课程及知识点坐标
            y3=y2_2+sprt
            y3_2=y3+s3
            #照片坐标
            y4=y3_2+sprt 
            if num>3:
                y4_2=y4+s4
            else:
                y4_2=y4+s4//2
            self.y4=y4
            self.y4_2=y4_2           
            #课堂情况（老师评论）坐标，变量
            y5=y4_2+sprt
            y5_2=y5+s5
            self.y5=y5
            self.y5_2=y5_2
            #积分坐标
            y_score=y5_2+sprt
            y_score_2=y_score+s_score
            self.y_score=y_score
            self.y_score_2=y_score_2
            #logo、二维码坐标
            y6=y_score_2+sprt
            y6_2=y6+s6    
            self.y6=y6
            self.y6_2=y6_2               

            if num==4:
                img = Image.new("RGB",(900,total_len),(255,255,255))
                draw=ImageDraw.Draw(img)
                draw.rectangle([(0,0),(900,y1_2)],fill=color['title_bg']) #乐高机器人课
                draw.rectangle([(0,y2),(900,y2_2)],fill=color['name']) # 姓名 年龄
                draw.rectangle([(0,y3),(900,y3_2)],fill=color['crs']) # 课程
                draw.rectangle([(0,y4),(900,y4_2)],fill=color['pics'])# 照片
                draw.rectangle([(0,y5),(300,y5+80)],fill=color['comments']) # 能力测评标题背景
                draw.rectangle([(0,y5+80),(900,y5_2)],fill=color['comments']) # 能力测评
                draw.rectangle([(0,y6),(900,y6_2)],fill=color['logo']) # logo


                draw.rectangle([(pic0[0][0],pic0[0][1]),(pic0[0][2],pic0[0][3])],fill=color['pics_box']) #相框_课程
                draw.rectangle([(pic1[0][0],pic1[0][1]),(pic1[0][2],pic1[0][3])],fill=color['pic01']) #相框_1
                draw.rectangle([(pic2[0][0],pic2[0][1]),(pic2[0][2],pic2[0][3])],fill=color['pic02']) #相框_2
                draw.rectangle([(pic3[0][0],pic3[0][1]),(pic3[0][2],pic3[0][3])],fill=color['pic03']) #相框_3
                draw.rectangle([(pic4[0][0],pic4[0][1]),(pic4[0][2],pic4[0][3])],fill=color['pic04']) #相框_4

                draw.rectangle([(0,y_score),(300,y_score+80)],fill=color['scores']) #积分标题背景
                draw.rectangle([(0,y_score+80),(900,y_score_2)],fill=color['scores']) #积分
                draw.rectangle([(20,y_score+80+20),(440,y_score_2-20)],fill=color['score_left']) #积分左
                draw.rectangle([(460,y_score+80+20),(880,y_score_2-20)],fill=color['score_right']) #积分右
            elif num==2:
                img = Image.new("RGB",(900,int(total_len)),(255,255,255))
                draw=ImageDraw.Draw(img)
                draw.rectangle([(0,0),(900,y1_2)],fill=color['title_bg']) #乐高机器人课
                draw.rectangle([(0,y2),(900,y2_2)],fill=color['name']) # 姓名 年龄
                draw.rectangle([(0,y3),(900,y3_2)],fill=color['crs']) # 课程
                draw.rectangle([(0,y4),(900,y4_2 )],fill=color['pics'])# 照片
                draw.rectangle([(0,y5),(300,y5+80)],fill=color['comments']) # 能力测评标题背景
                draw.rectangle([(0,y5+80),(900,y5_2)],fill=color['comments']) # 能力测评
                draw.rectangle([(0,y6),(900,y6_2)],fill=color['logo']) # logo


                draw.rectangle([(pic0[0][0],pic0[0][1]),(pic0[0][2],pic0[0][3])],fill=color['pics_box']) #相框_课程
                draw.rectangle([(pic1[0][0],pic1[0][1]),(pic1[0][2],pic1[0][3])],fill=color['pic01']) #相框_1
                draw.rectangle([(pic2[0][0],pic2[0][1]),(pic2[0][2],pic2[0][3])],fill=color['pic02']) #相框_2
                # draw.rectangle([(pic3[0][0],pic3[0][1]),(pic3[0][2],pic3[0][3])],fill=color['pic03']) #相框_3
                # draw.rectangle([(pic4[0][0],pic4[0][1]),(pic4[0][2],pic4[0][3])],fill=color['pic04']) #相框_4
                draw.rectangle([(0,y_score),(300,y_score+80)],fill=color['scores']) #积分标题背景
                draw.rectangle([(0,y_score+80),(900,y_score_2)],fill=color['scores']) #积分
                draw.rectangle([(20,y_score+80+20),(440,y_score_2-20)],fill=color['score_left']) #积分左
                draw.rectangle([(460,y_score+80+20),(880,y_score_2-20)],fill=color['score_right']) #积分右
            

            
            return img  
        
        def pick_pics(stdName):
            # print('stdname:',stdName)
            pic_title_addr=os.path.join(self.ConsDir,crs_nameInput,crs_name+'.jpg')  #课程的标题图
            
            ptn='-.*-'
            pics_for_crs=[]
            
            for root,dirs,files in os.walk(os.path.join(self.picStdDir,stdName)):   #学员的照片
                for file in files:
                    try:
                        if re.findall(ptn,file)[0][1:-1]==crs_nameInput:
                            PicTags=readConfig.code_to_str(iptcinfo3.IPTCInfo(os.path.join(self.picStdDir,stdName,file)))
                            # print('tags:',PicTags)
                            if '每周课程4+' in PicTags:
                                pics_for_crs.append(os.path.join(self.picStdDir,stdName,file))
                    except:
                        pass
            # print(pics_for_crs)
            # pics=pick_pics(stdName)            
            if len(pics_for_crs)>3:
                num=4
            else:
                num=2

            # print('number:',num)
            # print(pics_for_crs)
            pics_stds_addrs=random.sample(pics_for_crs,num)

            # print(pics_stds_addrs)
            sorted_pics_stds_addrs=sortPics(pics_stds_addrs)
            pics=[pic_title_addr]
            pics.extend(sorted_pics_stds_addrs)        
            return pics
        
        def putImg(img,stdName):
            print('    正在置入图片……',end='')
            pics=pick_pics(stdName)

            if self.bg_img_num>3:
                f_crs,f_01,f_02,f_03,f_04=pics

                pic_crs=self.pic_resize(Image.open(f_crs),picWid)
                pic_01=self.pic_resize(Image.open(f_01),picWid)
                pic_02=self.pic_resize(Image.open(f_02),picWid)
                pic_03=self.pic_resize(Image.open(f_03),picWid)
                pic_04=self.pic_resize(Image.open(f_04),picWid)
                _pic_logo=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超新logo.png').convert('RGBA')
                pic_logo=_pic_logo.resize((350,int(350/2.76)))
                r,g,b,a=pic_logo.split()

                _pic_qrcode=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超视频号二维码2.png')
                pic_qrcode=_pic_qrcode.resize((130,130))

                _pic_award=Image.open('I:\\大智小超\\公共素材\\图片类\\奖状.png')
                pic_award=_pic_award.resize((420,420*261//419))
                r_ad,g_ad,b_ad,a_ad=pic_award.split()
                
                logging.info('照片尺寸：'+','.join(self.sstr(pic_crs.size)))
                img.paste(pic_crs,(pic0[1][0],pic0[1][1]))
                img.paste(pic_01,(pic1[1][0],pic1[1][1]))
                img.paste(pic_02,(pic2[1][0],pic2[1][1]))
                img.paste(pic_03,(pic3[1][0],pic3[1][1]))
                img.paste(pic_04,(pic4[1][0],pic4[1][1])) 
                img.paste(pic_logo,(50,int(self.y6+self.s6/2-350/2.76/2)),mask=a)
                img.paste(pic_qrcode,(700,int(self.y6+self.s6/2-130/2))) 
                img.paste(pic_award,(20,self.y_score+100),mask=a_ad)   
            else:
                f_crs,f_01,f_02=pics

                pic_crs=self.pic_resize(Image.open(f_crs),picWid)
                pic_01=self.pic_resize(Image.open(f_01),picWid)
                pic_02=self.pic_resize(Image.open(f_02),picWid)

                _pic_logo=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超新logo.png').convert('RGBA')
                pic_logo=_pic_logo.resize((350,int(350/2.76)))
                r,g,b,a=pic_logo.split()

                _pic_qrcode=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超视频号二维码2.png')
                pic_qrcode=_pic_qrcode.resize((130,130))

                _pic_award=Image.open('I:\\大智小超\\公共素材\\图片类\\奖状.png')
                pic_award=_pic_award.resize((420,420*261//419))
                r_ad,g_ad,b_ad,a_ad=pic_award.split()
                
                logging.info('照片尺寸：'+','.join(self.sstr(pic_crs.size)))
                img.paste(pic_crs,(pic0[1][0],pic0[1][1]))
                img.paste(pic_01,(pic1[1][0],pic1[1][1]))
                img.paste(pic_02,(pic2[1][0],pic2[1][1]))     

                img.paste(pic_logo,(50,int(self.y6+self.s6/2-350/2.76/2)),mask=a)
                img.paste(pic_qrcode,(700,int(self.y6+self.s6/2-130/2)))
                img.paste(pic_award,(20,self.y_score+100),mask=a_ad)   

            print('完成')
            
        def expScript(name_input):
            stdname=name_input      
            nickname=teacherCmt[teacherCmt['学生姓名']==stdname]['昵称'].values.tolist()[0]      
            teacherCmtTxt=teacherCmt[teacherCmt['学生姓名']==stdname][str(dateInput)+'-'+crs_nameInput].values.tolist()[0]
            
            # if teacherCmtTxt=='-':
            #     teacherCmtTxt=teacherCmt[teacherCmt['学生姓名']=='通用评论'][str(dateInput)+'-'+crs_nameInput].values.tolist()[0]
            prsTxt=random.choice(self.PraiseTxt)
            if teacherCmtTxt!='-':
                script=stdname+'同学在“'+crs_info[0]+'”这节课中，'+crs_info[2]+'\n'+teacherCmtTxt
            else:
                # script=stdname+'同学在“'+crs_info[0]+'”这节课中，'+crs_info[2]+'\n'+prsTxt
                #如果该学生评论为空，则读取通用评论。
                script=stdname+'同学在“'+crs_info[0]+'”这节课中，'+crs_info[2]+'\n'+teacherCmt[teacherCmt['学生姓名']=='通用评论'][str(dateInput)+'-'+crs_nameInput].values.tolist()[0]


            if '#' in script:
                script=script.replace('#',stdname)
            if '$' in script:
                if nickname=='-':
                    nickname=stdname
                script=script.replace('$',nickname)

            return script                   

        def get_std_scores(date_input,crs_name,std_name,df_this_crs_score):   
            this_score=df_this_crs_score['this_score']      
            std_this_score=this_score[this_score['学生姓名']==std_name][str(date_input)+'-'+crs_name].tolist()[0]

            all_scores=df_this_crs_score['all_scores']
            df_std_all_scores=all_scores[all_scores['学生姓名']==std_name]
            crs_sc=df_std_all_scores['课堂总积分'].tolist()[0]
            vrfy_sc=df_std_all_scores['核销积分'].tolist()[0]
            remain_sc=df_std_all_scores['剩余积分'].tolist()[0]
            std_all_scores={
                'crs_sc':crs_sc,
                'vrfy_sc':vrfy_sc,
                'remain_sc':remain_sc
            }

            return {'std_this_score':std_this_score,'std_all_scores':std_all_scores}

        def putTxt(img,stdName,stdAge,KdgtName,ClassName):   
            print('    正在置入文本……',end='')
            color=color_list('202101')
            draw=ImageDraw.Draw(img)        
            
            draw.text((170,5), '科', fill = color['t_title_ke1'],font=self.fonts('汉仪超级战甲',75))  #大题目
            draw.text((270,5), '学', fill = color['t_title_xue'],font=self.fonts('汉仪超级战甲',75))  #大题目
            draw.text((370,5), '机', fill = color['t_title_ji'],font=self.fonts('汉仪超级战甲',75))  #大题目
            draw.text((470,5), '器', fill = color['t_title_qi'],font=self.fonts('汉仪超级战甲',75))  #大题目
            draw.text((570,5), '人', fill = color['t_title_ren'],font=self.fonts('汉仪超级战甲',75))  #大题目
            draw.text((670,5), '课', fill = color['t_title_ke4'],font=self.fonts('汉仪超级战甲',75))  #大题目
            draw.text((350,120), stdName, fill =color['t_name'] ,font=self.fonts('汉仪心海行楷w',60))  #姓名
            # draw.text((530,160), str(stdAge)+'岁', fill = '#6AB34A',font=self.fonts('微软雅黑',60))  #年龄    
            draw.text((280,200), KdgtName, fill = color['t_kdgtn'],font=self.fonts('杨任东竹石体',33))  #幼儿园
            draw.text((460,200), ClassName, fill = color['t_class'],font=self.fonts('杨任东竹石体',33))  #班级
    
            draw.text((50,290), '• '+crs_info[0]+' •', fill = color['t_crs'],font=self.fonts('华康海报体W12(p)',40))  #课程名称  
            draw.text((50,360), '难度：'+crs_info[3], fill = color['t_diff'],font=self.fonts('汉仪锐智w',25))  #难度
            draw.text((530,630), '使用教具：'+crs_info[4], fill = color['t_tool'],font=self.fonts('微软雅黑',22))  #教具
            
            self.put_txt_img(img,crs_info[1],450,[25,420],25,fill = color['t_knlg'],font_name='汉仪锐智w',font_size=28)  #知识点    
            
            date_txt='-'.join([str(dateInput)[0:4],str(dateInput)[4:6],str(dateInput)[6:]])
            draw.text((100,630), date_txt, fill = color['t_date'],font=self.fonts('汉仪心海行楷w',35))  #日期           

            #评语&积分
            script=expScript(stdName)
            std_scores=get_std_scores(date_input=dateInput,crs_name=crs_nameInput,std_name=stdName,df_this_crs_score=df_std_scores)
            std_this_score=std_scores['std_this_score']
            crs_total_sc=std_scores['std_all_scores']['crs_sc']
            remain_sc=std_scores['std_all_scores']['remain_sc']
            vrfy_sc=std_scores['std_all_scores']['vrfy_sc']
            script=script.replace('*',str(std_this_score))

            self.put_txt_img(img,script,780,[60,self.y5+115],20,fill = color['t_tch_cmt'],font_name='汉仪字酷堂经解楷体w',font_size=36,addSPC='add_2spaces') #老师评语
            draw.text((650,int(self.y5_2-80)), TeacherSig, fill = color['t_tch_sig'],font=self.fonts('汉仪字酷堂经解楷体w',45) )  #签名                    
            draw.text((500,int(self.y6+self.s6/2-35)), '长按二维码 → \n关注视频号 →', fill = color['t_bottom'],font=self.fonts('微软雅黑',30)) 
            draw.text((20,self.y5+5), '我的课堂情况', fill = color['t_tch_cmt_title'],font=self.fonts('汉仪糯米团',45))  #我的课堂情况
            draw.text((60,self.y_score+5), '我的积分', fill = color['t_scores_title'],font=self.fonts('汉仪糯米团',45))  #我的积分
            draw.text((120,self.y_score+120), '本节课积分', fill = color['t_scores_left'],font=self.fonts('汉仪糯米团',45))  #本节课积分

            draw.text((210,self.y_score+200), str(std_this_score)+'分', fill = color['t_scores_left'],font=self.fonts('汉仪糯米团',75))  #本节课积分
            draw.text((490,self.y_score+160),'可兑换积分：'+str(int(remain_sc))+' 分', fill = color['t_scores_right'],font=self.fonts('优设标题',35))  #可兑换积分
            draw.text((490,self.y_score+210), '累计积分：'+str(int(crs_total_sc))+' 分', fill = color['t_scores_right'],font=self.fonts('优设标题',35))  #累计总积分
            draw.text((490,self.y_score+260), '已兑换积分：'+str(int(vrfy_sc))+' 分', fill = color['t_scores_right'],font=self.fonts('优设标题',35))  #消耗积分
            
           
            print('完成')
            
        para=basic_para()        
        
        picWid,picX1,picX2,picX3,picX4,picY1,picY2,picY3,picY4,pic0,pic1,pic2,pic3,pic4= \
        para[0],para[1],para[2],para[3],para[4],para[5],para[6],para[7],para[8],para[9],para[10],para[11],para[12],para[13]
        
        INFO=read_excel()
        df_std_scores=read_scores(place_input=self.place_input,term=self.term,weekday=self.weekday)
        std_list,crs_info,teacherCmt=INFO['std_list'],INFO['crs_info'],INFO['tch_cmt']
    
        for std in std_list:
            print('正在处理 {} 的图片：'.format(std[3]))
            # img=basic_bg()
            KdgtName,ClassName,stdPY,stdName,stdAge=std[0],std[1],std[2],std[3],'-'
            totalTxt=expScript(stdName)
            self.allTxt,self.para_num=composing.split_txt_Chn_eng(wid=780,font_size=self.comment_font_size,txt_input=totalTxt,Indent='yes')



            self.bg_img_num=len(pick_pics(std[2]+std[3]))
            if self.bg_img_num>3:
                img=basic_bg(4)
            else:
                img=basic_bg(2)
            
            putImg(img,stdPY+stdName)
            putTxt(img,stdName,stdAge,KdgtName,ClassName)             
            
            print('正在保存 {} 的图片……'.format(std[3]),end='')
            if not os.path.exists(os.path.join(self.bg,str(dateInput)[0:4],str(dateInput)+'-'+crs_nameInput)):
                os.mkdir(os.path.join(self.bg,str(dateInput)[0:4],str(dateInput)+'-'+crs_nameInput))
            
            img.save(os.path.join(self.bg,str(dateInput)[0:4],str(dateInput)+'-'+crs_nameInput,std[2]+stdName+'-'+str(dateInput)+'-'+crs_nameInput+'.jpg'))
            print('完成')
            # img.show()


        print('\n全部完成,保存文件夹：{} 下面的学生文件名'.format(self.bg))
        
    
if __name__=='__main__':
    my=poster(weekday=4,term='2021春')
#     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw('L074小青蛙',20210422,TeacherSig='阿晓老师')