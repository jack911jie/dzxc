import sys
sys.path.append('i:/py/dzxc')
import composing
import os
import json
import time
import re
import random
import logging
import pandas as pd
from PIL import Image,ImageDraw,ImageFont
from tqdm import tqdm
import exifread

logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(funcName)s-%(lineno)d - %(message)s')
logger = logging.getLogger(__name__)

class poster:
    def __init__(self):
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
        self.crsStudent=config['学员签到表']
        self.picTitleDir=config['课程标题照片文件夹']
        self.picStdDir=config['学员照片文件夹']
        
        self.picWid=425 #默认照片富庶
        
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
        txt=composing.split_txt_Chn_eng(total_dis,font_size,t,Indent='yes')
        # font_sig = self.fonts('丁永康硬笔楷书',40)
        num=len(txt)   
        draw=ImageDraw.Draw(img)
   
        logging.info(txt)
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
                        draw.text((x+font_size*0.2,y), tt, fill = fill,font=fontInput) 
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
    
    def PosterDraw(self,crs_name,dateInput):
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
            for file in files:
                if file[-3:].lower()=='jpg':
                    with open(file,'rb') as fd:
                        tag=exifread.process_file(fd)
                        t=str(tag['EXIF DateTimeOriginal']).replace(':','-',2)
                        newfiles[t]=file
                        # print(t,file)
            newList=[]
            for i in sorted(newfiles):
                newList.append(newfiles[i])
            return newList
        
        def read_excel():
            print('正在读取学员和课程信息……',end='')
            df=pd.read_excel(self.crsList)        
            crs=df.loc[df['课程名称']==crs_name]   
            knowledge=list(crs['知识点'])
            script=list(crs['话术'])
            dif_level=list(crs['难度'])
            instrument=list(crs['教具'])
            crs_info=[crs_name,knowledge[0],script[0],dif_level[0],instrument[0]]            
            # print(crs_info)
            stars=crs_info[-1].replace('*','★')
            crs_info[-1]=stars 
            
            df_stdInfo=pd.read_excel(self.crsStudent,sheet_name='学生基本信息表')
            df_stdSig=pd.read_excel(self.crsStudent,sheet_name='学生上课签到表',skiprows=2)
            
            df_stdSig.rename(columns={'Unnamed: 0':'幼儿园','Unnamed: 1':'班级','Unnamed: 2':'姓名首拼','Unnamed: 3':'性别','Unnamed: 4':'学生姓名'},inplace=True)
            Students_sig=df_stdSig.loc[df_stdSig[crs_name]=='√'][['幼儿园','班级','姓名首拼','学生姓名']] #上课的学生名单            
            Students=pd.merge(Students_sig,df_stdInfo,on='学生姓名',how='left') #根据学生名单获取学生信息
            Students_List=Students.values.tolist()
            logging.info(Students_List)
            # logging.info('\n'.join(crs_info))   
            print('完成')
            return([Students_List,crs_info])               
        
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
        
        def basic_bg(num):
            s1=100
            s2=140
            s3=450
            s4=700
            s5=380
            s6=200
            sprt=5      
            total_len=s1+s2+s3+s4+s5+s6+sprt*5

            y1=s1

            y2=y1+sprt        
            y2_2=y2+s2

            y3=y2_2+sprt
            y3_2=y3+s3

            y4=y3_2+sprt 
            y4_2=y4+s4

            y5=y4_2+sprt
            y5_2=y5+s5

            y6=y5_2+sprt
            y6_2=y6+s6       

            if num==4:
                img = Image.new("RGB",(900,total_len),(255,255,255))
                draw=ImageDraw.Draw(img)
                draw.rectangle([(0,0),(900,y1)],fill='#FF9E11') #乐高机器人课
                draw.rectangle([(0,y2),(900,y2_2)],fill='#FFFFFF') # 姓名 年龄
                draw.rectangle([(0,y3),(900,y3_2)],fill='#00B578') # 课程
                draw.rectangle([(0,y4),(900,y4_2)],fill='#F7F7F7')# 照片
                draw.rectangle([(0,y5),(900,y5_2)],fill='#006DE3') # 能力测评
                draw.rectangle([(0,y6),(900,y6_2)],fill='#FFFFFF') # logo


                draw.rectangle([(pic0[0][0],pic0[0][1]),(pic0[0][2],pic0[0][3])],fill='#FFFFFF') #相框_课程
                draw.rectangle([(pic1[0][0],pic1[0][1]),(pic1[0][2],pic1[0][3])],fill='#FFFFFF') #相框_1
                draw.rectangle([(pic2[0][0],pic2[0][1]),(pic2[0][2],pic2[0][3])],fill='#FFFFFF') #相框_2
                draw.rectangle([(pic3[0][0],pic3[0][1]),(pic3[0][2],pic3[0][3])],fill='#FFFFFF') #相框_3
                draw.rectangle([(pic4[0][0],pic4[0][1]),(pic4[0][2],pic4[0][3])],fill='#FFFFFF') #相框_4
            elif num==2:
                img = Image.new("RGB",(900,int(total_len-s4/2)),(255,255,255))
                draw=ImageDraw.Draw(img)
                draw.rectangle([(0,0),(900,y1)],fill='#FF9E11') #乐高机器人课
                draw.rectangle([(0,y2),(900,y2_2)],fill='#FFFFFF') # 姓名 年龄
                draw.rectangle([(0,y3),(900,y3_2)],fill='#00B578') # 课程
                draw.rectangle([(0,y4),(900,y4_2-s4/2)],fill='#F7F7F7')# 照片
                draw.rectangle([(0,y5-s4/2),(900,y5_2-s4/2)],fill='#006DE3') # 能力测评
                draw.rectangle([(0,y6-s4/2),(900,y6_2-s4/2)],fill='#FFFFFF') # logo


                draw.rectangle([(pic0[0][0],pic0[0][1]),(pic0[0][2],pic0[0][3])],fill='#FFFFFF') #相框_课程
                draw.rectangle([(pic1[0][0],pic1[0][1]),(pic1[0][2],pic1[0][3])],fill='#FFFFFF') #相框_1
                draw.rectangle([(pic2[0][0],pic2[0][1]),(pic2[0][2],pic2[0][3])],fill='#FFFFFF') #相框_2
                # draw.rectangle([(pic3[0][0],pic3[0][1]),(pic3[0][2],pic3[0][3])],fill='#FFFFFF') #相框_3
                # draw.rectangle([(pic4[0][0],pic4[0][1]),(pic4[0][2],pic4[0][3])],fill='#FFFFFF') #相框_4

            
            return img  
        
        def pick_pics(stdName):
            # print('stdname:',stdName)
            pic_title_addr=os.path.join(self.picTitleDir,str(dateInput)+crs_name,crs_name+'.jpg')  #课程的标题图
            
            ptn='-.*-'
            pics_for_crs=[]

            
            for root,dirs,files in os.walk(os.path.join(self.picStdDir,stdName)):   #学员的照片
                for file in files:
                    try:
                        if re.findall(ptn,file)[0][1:-1]==crs_name:
                            pics_for_crs.append(os.path.join(self.picStdDir,stdName,file))
                    except:
                        pass
            
            # pics=pick_pics(stdName)
            if len(pics_for_crs)>3:
                num=4
            else:
                num=2

            pics_stds_addrs=random.sample(pics_for_crs,num)
            sorted_pics_stds_addrs=sortPics(pics_stds_addrs)
            pics=[pic_title_addr]
            pics.extend(sorted_pics_stds_addrs)
            
            logging.info(pics)
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
                img.paste(pic_logo,(50,1830),mask=a)
                img.paste(pic_qrcode,(700,1830))    
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
                img.paste(pic_logo,(50,1490),mask=a)
                img.paste(pic_qrcode,(700,1490))    

            print('完成')
            
        def expScript(name_input):
            name=name_input
            script=name+'在“'+crs_info[0]+'”这节课中，'+crs_info[2]
            return script                   

        def putTxt(img,stdName,stdAge,KdgtName,ClassName):   
            print('    正在置入文本……',end='')
            draw=ImageDraw.Draw(img)        
            
            draw.text((200,6), '科 学 机 器 人 课', fill = '#FFFFFF',font=self.fonts('汉仪超级战甲',75))  #大题目
            draw.text((350,120), stdName, fill = '#00b578',font=self.fonts('汉仪心海行楷w',60))  #姓名
            # draw.text((530,160), str(stdAge)+'岁', fill = '#6AB34A',font=self.fonts('微软雅黑',60))  #年龄    
            draw.text((280,200), KdgtName, fill = '#00b578',font=self.fonts('杨任东竹石体',33))  #幼儿园
            draw.text((460,200), ClassName, fill = '#00b578',font=self.fonts('杨任东竹石体',33))  #班级
    
            draw.text((50,290), '• '+crs_info[0], fill = '#ffffff',font=self.fonts('华康海报体W12(p)',40))  #课程名称  
            draw.text((50,360), '难度：'+crs_info[3], fill = '#ffffff',font=self.fonts('汉仪锐智w',22))  #难度
            draw.text((530,630), '使用教具：'+crs_info[4], fill = '#ffffff',font=self.fonts('微软雅黑',22))  #教具
            
            self.put_txt_img(img,crs_info[1],450,[8,420],25,fill = '#ffffff',font_name='汉仪锐智w',font_size=28)  #知识点    
            
            date_txt='-'.join([str(dateInput)[0:4],str(dateInput)[4:6],str(dateInput)[6:]])
            draw.text((100,630), date_txt, fill = '#ffffff',font=self.fonts('汉仪心海行楷w',35))  #日期
            
            # draw.text((50,1490), '能力测评', fill = '#6AB34A',font=font_2)  #能力测评
            # draw.text((50,1560), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            # draw.text((50,1610), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            # draw.text((50,1660), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            # draw.text((50,1710), 'XX力', fill = '#6AB34A',font=font_3)  #XX力
            script=expScript(stdName)
            if self.bg_img_num>3:
                self.put_txt_img(img,script,780,[60,1440],20,fill = '#ffffff',font_name='丁永康硬笔楷书',font_size=36,addSPC='add_2spaces') #老师评语

                draw.text((650,1730), '阿晓老师', fill = '#ffffff',font=self.fonts('丁永康硬笔楷书',45) )  #签名    
                
                draw.text((500,1860), '长按二维码 → \n关注视频号 →', fill = '#656564',font=self.fonts('微软雅黑',30))  
            else:
                self.put_txt_img(img,script,780,[60,1100],20,fill = '#ffffff',font_name='丁永康硬笔楷书',font_size=36,addSPC='add_2spaces') #老师评语

                draw.text((650,1390), '阿晓老师', fill = '#ffffff',font=self.fonts('丁永康硬笔楷书',45) )  #签名    
                
                draw.text((500,1520), '长按二维码 → \n关注视频号 →', fill = '#656564',font=self.fonts('微软雅黑',30)) 
                
            print('完成')
            
         para=basic_para()        
        
        picWid,picX1,picX2,picX3,picX4,picY1,picY2,picY3,picY4,pic0,pic1,pic2,pic3,pic4= \
        para[0],para[1],para[2],para[3],para[4],para[5],para[6],para[7],para[8],para[9],para[10],para[11],para[12],para[13]
        
        INFO=read_excel()
        std_list,crs_info=INFO[0],INFO[1]
        for std in std_list:
            print('正在处理 {} 的图片：'.format(std[3]))
            # img=basic_bg()
            self.bg_img_num=len(pick_pics(std[2]+std[3]))
            if self.bg_img_num>3:
                img=basic_bg(4)
            else:
                img=basic_bg(2)
            KdgtName,ClassName,stdPY,stdName,stdAge=std[0],std[1],std[2],std[3],'-'
            putImg(img,stdPY+stdName)
            putTxt(img,stdName,stdAge,KdgtName,ClassName)             
            
            print('正在保存 {} 的图片……'.format(std[3]),end='')
            if not os.path.exists(os.path.join(self.bg,str(dateInput)+crs_name)):
                os.mkdir(os.path.join(self.bg,str(dateInput)+crs_name))
            
            img.save(os.path.join(self.bg,str(dateInput)+crs_name,stdName+'-'+str(dateInput)+'-'+crs_name+'.jpg'))
            print('完成')
#             img.show()


        print('\n全部完成,保存文件夹：{} 下面的学生文件名'.format(self.bg))
        
       
if __name__=='__main__':
    my=poster()
#     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw('夏天的手摇风扇',20201013)    