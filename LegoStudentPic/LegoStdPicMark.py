import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'../../dzxc/module'))
import composing as paraFormat
import re
import json
import pandas as pd
import iptcinfo3
from PIL import Image,ImageFont,ImageDraw
from tqdm import tqdm
import logging
iptcinfo_logger=logging.getLogger('iptcinfo')
iptcinfo_logger.setLevel(logging.ERROR)


class pics:
    def __init__(self):
        print('正在初始化参数……',end='')
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'StudentsPicConfig.txt'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
            config=json.loads(_line)
        
        self.publicPicDir=config['公共图片']
        self.StdInfoDir=config['2020乐高课程签到表文件夹']
        # self.StdInfoDir=config['2019科学课签到表文件夹']
        
        self.CrsInfoDir=config['课程信息表']
        self.totalPics=config['照片总文件夹']
        self.stdPicsDir=config['照片文件夹']
        # print(self.publicPicDir,self.CrsInfo)
        print('完成')
        

    def putCover(self,height=2250,weekday=2):
        def read_excel():
            crsFile=['课程信息表.xlsx','课程信息']
            if weekday==2:
                stdFile=['2020乐高课程签到表（周二）.xlsx','学生上课签到表']
            elif weekday==6:
                stdFile=['2020乐高课程签到表（周六）.xlsx','学生上课签到表']
            # stdFile=['2019科学实验课学员档案2.xlsx','学员名单']
            crs=pd.read_excel(os.path.join(self.CrsInfoDir,crsFile[0]),skiprows=0,sheet_name=crsFile[1])
            stds=pd.read_excel(os.path.join(self.StdInfoDir,stdFile[0]),skiprows=2,sheet_name=stdFile[1])
            stds.rename(columns={'Unnamed: 0':'机构','Unnamed: 1':'班级','Unnamed: 2':'姓名首拼','Unnamed: 3':'性别','Unnamed: 4':'ID','Unnamed: 5':'学生姓名','Unnamed: 6':'已上课数量'},inplace=True)
            # std=stds[stds['学生姓名']==stdName]
            # std_basic=std[['姓名首拼','学生姓名']]
            # std_crs=std[std.iloc[:,:]=='√'].dropna(axis=1)
            # std_res=pd.concat([std_basic,std_crs],axis=1)
            # print(crs)
            return [crs,stds]

        def read_pics_new():
            print('正在读取照片……',end='')
            lst=read_excel()
            crs,stds=lst[0],lst[1]
            stdList=stds['学生姓名'].tolist()
            ptn_pic_src=re.compile(r'[0-9]{8}\-[a-zA-Z].*')
            ptn_std_name=re.compile(r'^[a-zA-Z]+[\u4e00-\u9fa5]+')

            infos=[]
            for fileName in os.listdir(self.totalPics):
                fn=fileName.split('-')
                crsName=fn[1][4:]
                crsCode=fn[1][0:4]
                real_addr=os.path.join(self.totalPics,fileName)
                tag=self.code_to_str(iptcinfo3.IPTCInfo(real_addr))
                if len(tag)>0:
                    for _tag in tag:        
                        # print(_tag)     
                        _tag.strip()
                        _tag=_tag.replace(' ','')           
                        if ptn_std_name.match(_tag):
                            _tag=re.findall(r'[\u4e00-\u9fa5]+',_tag)[0] 
                        # print('77 _tag:',_tag)

                        if _tag in stdList:
                            # print('80_tag:',_tag)
                            std_name=_tag
                            knlg=crs[crs['课程编号']==crsCode]['知识点'].tolist()[0]
                            infos.append([real_addr,std_name,crsName,knlg,fileName])
            print('完成')
            # print('77infos:',infos)
            return infos

        def ResizeCrop(pic,h_min=2250,crop='yes',bigger='yes'):
            img=Image.open(pic)
            w,h=img.size

            if h>=2250:
                if w/h!=0.75:
                    lth=int(h_min*w/h)
                    pp=img.resize((lth,h_min))
                    if crop=='yes':
                        pp_crop=pp.crop((lth-h_min*4/3,0,lth,h_min)) 
                        return pp_crop
                    else:
                        return pp
                else:
                    return img
            else:
                if bigger=='yes':
                    lth=int(h_min*w/h)
                    ppp=img.resize((lth,h_min))
                    if crop=='yes':
                        ppp_crop=ppp.crop((lth-h_min*4/3,0,lth,h_min))         
                        return ppp_crop
                    else:
                        return ppp
                else:
                    return 'picture is too small' 

        def ratio(p,rate):
            k=int(p*rate)
            return k

        def draw(img,w,h,txt):           
            

            draw=ImageDraw.Draw(img)
            draw.rectangle([(0,int(img.size[1]-h)),(w,img.size[1])],fill='#eae8e8') #背景
            r=img.size[1]/3024
    
            # print(img.size,w,h)
            title='\n'.join(paraFormat.split_txt_Chn_eng(ratio(360,r),ratio(90,r),txt[1])[0][0])
            _pic_logo=Image.open('I:\\大智小超\\公共素材\\图片类\\logoForPic.png').convert('RGBA')
            _pic_logo2=Image.open('I:\\大智小超\\公共素材\\图片类\\logoDZXC.png').convert('RGBA')
            pic_logo=_pic_logo.resize((ratio(450,r),ratio(450/1.7616,r)))
            pic_logo2=_pic_logo2.resize((ratio(350,r),ratio(350*46/36,r)))
            red,g,b,a=pic_logo.split()
            red2,g2,b2,a2=pic_logo2.split()
            _pic_qrcode=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超视频号二维码2.png')
            # print(ratio(350,r),ratio(350,r))
            pic_qrcode=_pic_qrcode.resize((ratio(350,r),ratio(350,r)))
            img.paste(pic_logo,(ratio(50,r),ratio(100,r)),mask=a)
            img.paste(pic_logo2,(ratio(120,r),ratio(2500,r)),mask=a2)
            img.paste(pic_qrcode,(ratio(3560,r),ratio(2500,r)))    
            draw.text((ratio(3560,r),ratio(2880,r)), '微信扫码关注视频号', fill = '#000000',font=paraFormat.fonts('微软雅黑',ratio(40,r)))

            partTitle=ratio(1000,r)
            partKnlg=ratio(1500,r)
            titleSize=ratio(300,r)
            knlgSize=ratio(80,r)
            dateSize=ratio(65,r)
            xTitle=ratio(700,r)
            
            # txt[1]='恐龙'
            # print('before  ',titleSize)
            while len(txt[1])*titleSize>partTitle:
                titleSize=titleSize-2
            # print('after  ',titleSize)
            xDate=int(len(txt[1])*titleSize/2+xTitle)-int(paraFormat.char_len(txt[2])*dateSize/2)-20


            draw.text((xTitle,int(ratio(2400,r)+0.8*(ratio(300,r)-titleSize))), txt[1], fill = '#2A68B1',font=paraFormat.fonts('优设标题',titleSize))  #课程题目  单个汉字的上方会有空间，空间大小与字体成正比，所以y坐标要根据字体大小改变。
            draw.text((xDate,ratio(2800,r)), txt[2], fill = '#2A68B1',font=paraFormat.fonts('微软雅黑',dateSize))  #日期，坐标根据课程题目调整，居中对齐
            # draw.text((1800,2500), txt[3], fill = '#2A68B1',font=paraFormat.fonts('杨任东竹石体',140))  #知识点
            draw.text((ratio(1900,r),ratio(2450,r)), '课程知识点', fill = '#2A68B1',font=paraFormat.fonts('汉仪字酷堂义山楷w',80))  #知识点
            draw.rectangle((ratio(1900,r),ratio(2560,r),ratio(2400,r),ratio(2565,r)),'#2A68B1')

            txt3Dot=re.sub('\d.','· ',txt[3]) #将数字替换成点
            # print('153:',isinstance(txt3Dot,str))
            paraFormat.put_txt_img(draw,txt3Dot,partKnlg,[ratio(1850,r),ratio(2620,r)],25,'#2A68B1','楷体',knlgSize) #知识点，可设置行间距

        def putCoverToPics():
            infos=read_pics_new()
            # print(infos)
            # infos=[infos[2]]            
            if infos:
                print('正在写入标注信息并按姓名保存到文件夹')
                smallpics=[]
                for info in tqdm(infos):
                    date_crs=info[4].split('-')[0][0:4]+'-'+info[4].split('-')[0][4:6]+'-'+info[4].split('-')[0][6:]
                    # img=Image.open(info[0])
                    img=ResizeCrop(info[0],h_min=height,bigger='no')
                    if not isinstance(img,str):
                        bg_h,bg_w=int(img.size[1]*0.2018),img.size[0]
                        txt_write=[info[0],info[2],date_crs,info[3]]
                        # print('173 txt_write:',txt_write)
                        draw(img,bg_w,bg_h,txt_write)
                        saveDir=os.path.join(self.stdPicsDir,info[1])
                        saveName=os.path.join(saveDir,info[4])
                        if not os.path.exists(saveDir):
                            # print(saveName)
                            os.mkdir(saveDir)
                            img.save(saveName,quality=95,subsampling=0) #subsampling参数：子采样，通过实现色度信息的分辨率低于亮度信息来对图像进行编码的实践。可能的子采样值是0,1和2。
                        else:
                            img.save(saveName,quality=95,subsampling=0)
                            # print(saveName)
                        # print('生成的照片所在文件夹：{}'.format(saveDir))
                    else:
                        smallpics.append(info[4])

                if smallpics:
                    msg='完成 {}/{} 个文件。{}个文件大小，未完成：'.format(len(infos)-len(smallpics),len(infos),len(smallpics))+', '.join(smallpics)+'   too small.'
                    print(msg)
                print('完成')
            else:
                print('无合适的照片')

            
                


        putCoverToPics()

    def code_to_str(self,ss):
        s=ss['keywords']
        if isinstance(s,list):
            out=[]
            for i in s:
                out.append(i.decode('utf-8'))
        else:
            out=[ss.decode('utf-8')]
    
        return out

if __name__=='__main__':
    pic=pics()
    pic.putCover(height=2250,weekday=6)