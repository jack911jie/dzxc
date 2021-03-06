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
        # self.eachStd=config['个别学员评语表']
        # 个别学员评语表":"E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\学员课堂学习情况反馈表.xlsx",
        cmt_table_dir=config['老师评语表文件夹']
        self.eachStd=os.path.join(cmt_table_dir,place_input,'每周课程反馈',term+'-学员课堂学习情况反馈表.xlsx')
        
        self.picTitleDir=config['课程标题照片文件夹']
        self.picStdDir=config['学员照片文件夹']
        self.ConsDir=config['乐高图纸文件夹']
        self.std_sig_dir=config['学员签到表文件夹']

        # if weekday==2:
        #     self.crsStudent=config['学员签到表w2']
        # elif weekday==6:
        #     self.crsStudent=config['学员签到表w6']
        wd=days_calculate.num_to_ch(str(weekday))
        self.crsStudent=os.path.join(self.std_sig_dir,place_input,term+'-乐高课程签到表（周'+wd+'）.xlsx')

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
            '杨任东竹石体':'j:\\fonts\\yangrendongzhushi-Regular.ttf'            
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
            infos=WashData.comments_after_class(crs_name_input=crs_nameInput,weekday=self.weekday, \
                                                crs_list=self.crsList,crs_student=self.crsStudent, \
                                                tch_cmt=self.eachStd)
            print('完成')
            return infos         
        
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
                't_name':'#00b578',#姓名
                't_kdgtn':'#00b578',#幼儿园
                't_class':'#00b578',#班级
                't_crs':'#FFE340',#课程名称
                't_diff':'#ffffff',#难度
                't_tool':'#ffffff',#教具
                't_knlg':'#ffffff',#知识点
                't_date':'#ffffff',#日期
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
                't_tch_cmt':'#795022',#老师评语
                't_tch_sig':'#795022',#签名
                't_bottom':'#656564' #二维码文字
            }

            out=eval('colors_'+list_num)

            return out
        
        def basic_bg(num):
            color=color_list('202101')
            s1=100
            s2=140
            s3=450
            s4=700
            self.s4=s4
            s5=(self.comment_font_size+self.comment_dis_line)*self.para_num+120
            self.s5=s5
            s6=200
            self.s6=s6
            sprt=5      
            total_len=s1+s2+s3+s4+s5+s6+sprt*5

            y1=s1

            y2=y1+sprt        
            y2_2=y2+s2

            y3=y2_2+sprt
            y3_2=y3+s3

            y4=y3_2+sprt 
            y4_2=y4+s4
            self.y4=y4
            self.y4_2=y4_2

            y5=y4_2+sprt
            y5_2=y5+s5
            self.y5=y5
            self.y5_2=y5_2

            y6=y5_2+sprt
            y6_2=y6+s6    
            self.y6=y6
            self.y6_2=y6_2   

            

            if num==4:
                img = Image.new("RGB",(900,total_len),(255,255,255))
                draw=ImageDraw.Draw(img)
                draw.rectangle([(0,0),(900,y1)],fill=color['title_bg']) #乐高机器人课
                draw.rectangle([(0,y2),(900,y2_2)],fill=color['name']) # 姓名 年龄
                draw.rectangle([(0,y3),(900,y3_2)],fill=color['crs']) # 课程
                draw.rectangle([(0,y4),(900,y4_2)],fill=color['pics'])# 照片
                draw.rectangle([(0,y5),(900,y5_2)],fill=color['comments']) # 能力测评
                draw.rectangle([(0,y6),(900,y6_2)],fill=color['logo']) # logo


                draw.rectangle([(pic0[0][0],pic0[0][1]),(pic0[0][2],pic0[0][3])],fill=color['pics_box']) #相框_课程
                draw.rectangle([(pic1[0][0],pic1[0][1]),(pic1[0][2],pic1[0][3])],fill=color['pic01']) #相框_1
                draw.rectangle([(pic2[0][0],pic2[0][1]),(pic2[0][2],pic2[0][3])],fill=color['pic02']) #相框_2
                draw.rectangle([(pic3[0][0],pic3[0][1]),(pic3[0][2],pic3[0][3])],fill=color['pic03']) #相框_3
                draw.rectangle([(pic4[0][0],pic4[0][1]),(pic4[0][2],pic4[0][3])],fill=color['pic04']) #相框_4
            elif num==2:
                img = Image.new("RGB",(900,int(total_len-s4/2)),(255,255,255))
                draw=ImageDraw.Draw(img)
                draw.rectangle([(0,0),(900,y1)],fill=color['title_bg']) #乐高机器人课
                draw.rectangle([(0,y2),(900,y2_2)],fill=color['name']) # 姓名 年龄
                draw.rectangle([(0,y3),(900,y3_2)],fill=color['crs']) # 课程
                draw.rectangle([(0,y4),(900,y4_2-s4/2)],fill=color['pics'])# 照片
                draw.rectangle([(0,y5-s4/2),(900,y5_2-s4/2)],fill=color['comments']) # 能力测评
                draw.rectangle([(0,y6-s4/2),(900,y6_2-s4/2)],fill=color['logo']) # logo


                draw.rectangle([(pic0[0][0],pic0[0][1]),(pic0[0][2],pic0[0][3])],fill=color['pics_box']) #相框_课程
                draw.rectangle([(pic1[0][0],pic1[0][1]),(pic1[0][2],pic1[0][3])],fill=color['pic01']) #相框_1
                draw.rectangle([(pic2[0][0],pic2[0][1]),(pic2[0][2],pic2[0][3])],fill=color['pic02']) #相框_2
                # draw.rectangle([(pic3[0][0],pic3[0][1]),(pic3[0][2],pic3[0][3])],fill=color['pic03']) #相框_3
                # draw.rectangle([(pic4[0][0],pic4[0][1]),(pic4[0][2],pic4[0][3])],fill=color['pic04']) #相框_4

            
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
                
                logging.info('照片尺寸：'+','.join(self.sstr(pic_crs.size)))
                img.paste(pic_crs,(pic0[1][0],pic0[1][1]))
                img.paste(pic_01,(pic1[1][0],pic1[1][1]))
                img.paste(pic_02,(pic2[1][0],pic2[1][1]))
                img.paste(pic_03,(pic3[1][0],pic3[1][1]))
                img.paste(pic_04,(pic4[1][0],pic4[1][1])) 
                img.paste(pic_logo,(50,int(self.y5_2+self.s6/2-350/2.76/2)),mask=a)
                img.paste(pic_qrcode,(700,int(self.y5_2+self.s6/2-130/2)))    
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
                
                logging.info('照片尺寸：'+','.join(self.sstr(pic_crs.size)))
                img.paste(pic_crs,(pic0[1][0],pic0[1][1]))
                img.paste(pic_01,(pic1[1][0],pic1[1][1]))
                img.paste(pic_02,(pic2[1][0],pic2[1][1]))

                img.paste(pic_logo,(50,int(self.y5_2-self.s4/2)+int(self.s6/2-350/2.76/2)),mask=a)
                img.paste(pic_qrcode,(700,int(self.y5_2-self.s4/2)+int(self.s6/2-130/2)))    

            print('完成')
            
        def expScript(name_input):
            stdname=name_input      
            nickname=teacherCmt[teacherCmt['学生姓名']==stdname]['昵称'].values.tolist()[0]      
            teacherCmtTxt=teacherCmt[teacherCmt['学生姓名']==stdname][str(dateInput)+'-'+crs_nameInput].values.tolist()[0]
            prsTxt=random.choice(self.PraiseTxt)
            if teacherCmtTxt!='-':
                script=stdname+'在“'+crs_info[0]+'”这节课中，'+crs_info[2]+'\n'+teacherCmtTxt
            else:
                script=stdname+'在“'+crs_info[0]+'”这节课中，'+crs_info[2]+'\n'+prsTxt

            if '#' in script:
                script=script.replace('#',stdname)
            if '$' in script:
                if nickname=='-':
                    nickname=stdname
                script=script.replace('$',nickname)

            return script                   

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
            
            # draw.text((50,1490), '能力测评', fill = '#6AB34A',font=font_2)  #能力测评
            # draw.text((50,1560), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            # draw.text((50,1610), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            # draw.text((50,1660), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            # draw.text((50,1710), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            script=expScript(stdName)
            if self.bg_img_num>3:
                self.put_txt_img(img,script,780,[60,1440],20,fill = color['t_tch_cmt'],font_name='丁永康硬笔楷书',font_size=36,addSPC='add_2spaces') #老师评语

                draw.text((650,int(self.y5_2-45*2+45/2)), TeacherSig, fill = color['t_tch_sig'],font=self.fonts('丁永康硬笔楷书',45) )  #签名    
                
                draw.text((500,int(self.y5_2+self.s6/2-30)), '长按二维码 → \n关注视频号 →', fill = color['t_bottom'],font=self.fonts('微软雅黑',30))  
            else:
                self.put_txt_img(img,script,780,[60,1100],20,fill = color['t_tch_cmt'],font_name='丁永康硬笔楷书',font_size=36,addSPC='add_2spaces') #老师评语

                draw.text((650,int(self.y5_2-self.s4/2-45*2+45/2)), TeacherSig, fill = color['t_tch_sig'],font=self.fonts('丁永康硬笔楷书',45) )  #签名    
                
                draw.text((500,int(self.y5_2-self.s4/2)+int(self.s6/2-30)), '长按二维码 → \n关注视频号 →', fill = color['t_bottom'],font=self.fonts('微软雅黑',30)) 
                
            print('完成')
            
        para=basic_para()        
        
        picWid,picX1,picX2,picX3,picX4,picY1,picY2,picY3,picY4,pic0,pic1,pic2,pic3,pic4= \
        para[0],para[1],para[2],para[3],para[4],para[5],para[6],para[7],para[8],para[9],para[10],para[11],para[12],para[13]
        
        INFO=read_excel()
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
            if not os.path.exists(os.path.join(self.bg,str(dateInput)+'-'+crs_nameInput)):
                os.mkdir(os.path.join(self.bg,str(dateInput)+'-'+crs_nameInput))
            
            img.save(os.path.join(self.bg,str(dateInput)+'-'+crs_nameInput,std[2]+stdName+'-'+str(dateInput)+'-'+crs_nameInput+'.jpg'))
            print('完成')
            # img.show()


        print('\n全部完成,保存文件夹：{} 下面的学生文件名'.format(self.bg))
        
    
if __name__=='__main__':
    my=poster(weekday=6,term='2021春')
#     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw('L068厉害的投掷器',20210320,TeacherSig='阿晓老师')