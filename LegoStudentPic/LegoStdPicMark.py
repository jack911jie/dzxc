import os
import sys
# sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'../../dzxc/module'))
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),'modules'))
import readConfig
from composing import TxtFormat
from pics_modify import Shape
import days_calculate
import LegoStudentPicDistribute
import re
from datetime import datetime
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
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'LegoStudentPic.config'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
            config=json.loads(_line)
        
        self.publicPicDir=config['公共图片']
        self.StdInfoDir=config['科学机器人课程签到表文件夹']
        # self.StdInfoDir=config['2019科学课签到表文件夹']
        
        self.CrsInfoDir=config['课程信息表']
        self.totalPics=config['照片总文件夹']
        self.stdPicsDir=config['照片文件夹']
        self.after_class_dir=config['课后照片及反馈文件夹']
        # print(self.publicPicDir,self.CrsInfo)
        print('完成')
        

    def putCover(self,height=2250,term='2020秋',crop='yes',bigger='yes',weekday=2,savemode='all_together'): #savemode参数：all_together, individual
        def read_excel():
            crsFile=['课程信息表.xlsx','课程信息']
            # if weekday==2:
            #     stdFile=['2020乐高课程签到表（周二）.xlsx','学生上课签到表']
            # elif weekday==6:
            #     stdFile=['2020乐高课程签到表（周六）.xlsx','学生上课签到表']
            wd=days_calculate.num_to_ch(weekday)
            stdFile=[term+'-'+'学生信息表（周'+wd+'）.xlsx','学生上课签到表']
            # stdFile=['2019科学实验课学员档案2.xlsx','学员名单']
            crs=pd.read_excel(os.path.join(self.CrsInfoDir,crsFile[0]),skiprows=0,sheet_name=crsFile[1])
            stds=pd.read_excel(os.path.join(self.StdInfoDir,term[0:4],stdFile[0]),skiprows=1,sheet_name=stdFile[1])
            stds.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名', \
                                'Unnamed: 5':'昵称','Unnamed: 6':'性别','Unnamed: 7':'上期课时结余','Unnamed: 8':'购买课时','Unnamed: 9':'目前剩余课时', \
                                'Unnamed: 10':'上课数量统计汇总'},inplace=True)
            # print(stds)
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
            total_pics_dir=os.path.join(self.totalPics,term,term+'-每周课程16')
            # if not os.path.exists(total_pics_dir):
            #     os.makedirs(total_pics_dir)

            for fileName in os.listdir(total_pics_dir):
                if fileName[-3:].lower()=='jpg' or fileName[-4:].lower()=='jpeg':
                    fn=fileName.split('-')
                    crsName=fn[1][4:]
                    crsCode=fn[1][0:4]

                    # print(crsCode)
                    real_addr=os.path.join(total_pics_dir,fileName)
                    tag=self.code_to_str(iptcinfo3.IPTCInfo(real_addr))
                    if len(tag)>0:
                        for _tag in tag:        
                            # print(_tag)     
                            _tag.strip()
                            _tag=_tag.replace(' ','')           
                            if ptn_std_name.match(_tag):
                                _tag_py=re.findall(r'[a-zA-Z]+',_tag)[0] 
                                _tag_zh=re.findall(r'[\u4e00-\u9fa5]+',_tag)[0] 
                                
                            # print('77 _tag:',_tag)

                                if _tag_zh in stdList:
                                    # print('80_tag:',_tag)
                                    std_name=_tag_zh
                                    std_py=_tag_py
                                    knlg=crs[crs['课程编号']==crsCode]['知识点'].tolist()[0]
                                    infos.append([real_addr,std_name,crsName,knlg,fileName,std_py])
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

            rct=Image.new('RGBA',(w,h),(255,255,255,190))

            draw=ImageDraw.Draw(img)
            img.paste(rct,(0,int(img.size[1]-h)),mask=rct)

            # draw.rectangle([(0,int(img.size[1]-h)),(w,img.size[1])],fill='#eae8e8') #背景
            r=img.size[1]/3024
    
            # print(img.size,w,h)
            title='\n'.join(TxtFormat().split_txt_Chn_eng(ratio(360,r),ratio(90,r),txt[1])['txt'][0])
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
            draw.text((ratio(3560,r),ratio(2880,r)), '微信扫码关注视频号', fill = '#000000',font=TxtFormat().fonts('微软雅黑',ratio(40,r)))

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
            xDate=int(len(txt[1])*titleSize/2+xTitle)-int(TxtFormat().char_len(txt[2])*dateSize/2)-20


            draw.text((xTitle,int(ratio(2400,r)+0.8*(ratio(300,r)-titleSize))), txt[1], fill = '#2A68B1',font=TxtFormat().fonts('优设标题',titleSize))  #课程题目  单个汉字的上方会有空间，空间大小与字体成正比，所以y坐标要根据字体大小改变。
            draw.text((xDate,ratio(2800,r)), txt[2], fill = '#2A68B1',font=TxtFormat().fonts('微软雅黑',dateSize))  #日期，坐标根据课程题目调整，居中对齐
            # draw.text((1800,2500), txt[3], fill = '#2A68B1',font=TxtFormat().fonts('杨任东竹石体',140))  #知识点
            draw.text((ratio(1900,r),ratio(2450,r)), '课程知识点', fill = '#2A68B1',font=TxtFormat().fonts('汉仪字酷堂义山楷w',80))  #知识点

            draw.rectangle((ratio(1900,r),ratio(2560,r),ratio(2400,r),ratio(2565,r)),'#2A68B1')


            txt3Dot=re.sub('\d.','· ',txt[3]) #将数字替换成点
            # print('153:',isinstance(txt3Dot,str))
            TxtFormat().put_txt_img(draw,txt3Dot,partKnlg,[ratio(1850,r),ratio(2620,r)],25,'#2A68B1','楷体',knlgSize) #知识点，可设置行间距

        def putCoverToPics():
            infos=read_pics_new()

            std_name_count_list=[]
            for info_item in infos:
                std_name_count_list.append(info_item[5]+info_item[1])

            #计算张数
            std_name_count={}
            log_txt=[]
            for std_name in std_name_count_list:
                std_name_count[std_name]=std_name_count_list.count(std_name)

            #生成张数清单文本
            sort_list=sorted(list(std_name_count),key=str.lower)
            for n_std,std_name in enumerate(sort_list):
                if std_name not in log_txt:
                    log_txt.append(str(n_std+1).zfill(2)+'-'+std_name+': '+str(std_name_count[std_name])+'张')
            log_txt.append('共'+str(sum(std_name_count.values()))+'张')
      
            if infos:
                print('正在写入标注信息并按姓名保存到文件夹')
                smallpics=[]
                # log_txt=[]
                for info in tqdm(infos):
                    date_crs=info[4].split('-')[0][0:4]+'-'+info[4].split('-')[0][4:6]+'-'+info[4].split('-')[0][6:]
                    # img=Image.open(info[0])
                    img=ResizeCrop(info[0],h_min=height,crop=crop,bigger=bigger)
                    if not isinstance(img,str):
                        bg_h,bg_w=int(img.size[1]*0.2018),img.size[0]
                        txt_write=[info[0],info[2],date_crs,info[3]]

                        # print('173 txt_write:',txt_write)
                        draw(img,bg_w,bg_h,txt_write)
                        if savemode=='individual':
                            saveDir=os.path.join(self.stdPicsDir,term,'冲印版',str(weekday).zfill(2)+info[5]+info[1]+'-'+str(std_name_count[info[5]+info[1]])+'张')
                            saveName=os.path.join(saveDir,info[4])
                        elif savemode=='all_together':
                            saveDir=os.path.join(self.stdPicsDir,term,'冲印版','所有')
                            saveName=os.path.join(saveDir,info[5]+info[1]+'-'+info[4])


                        if not os.path.exists(saveDir):
                            # print(saveName)
                            os.makedirs(saveDir)
                        img.save(saveName,quality=95,subsampling=0) #subsampling参数：子采样，通过实现色度信息的分辨率低于亮度信息来对图像进行编码的实践。可能的子采样值是0,1和2。

                            # print(saveName)
                        # print('生成的照片所在文件夹：{}'.format(saveDir))
                        # print('测试保存:',saveName)
                    else:
                        smallpics.append(info[4])

                txt_log='\n'.join(log_txt)
                with open(os.path.join(saveDir,'照片张数清单.txt'),'w') as write_fn:
                    write_fn.write(txt_log)

                if smallpics:
                    msg='完成 {}/{} 个文件。{}个文件太小，未完成：'.format(len(infos)-len(smallpics),len(infos),len(smallpics))+', '.join(smallpics)+'   too small.'
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

class SimpleMark:
    def __init__(self,place_input='001-超智幼儿园'):
        config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'LegoStdPicMark.config'))
        self.std_pic_dir=config['乐高学员文件夹']
        self.std_pic_dir= self.std_pic_dir.replace('$',place_input)
        self.place_input=place_input
        self.public_pic_dir=config['公共图片']
        self.save_dir=config['生成照片文件夹']
        self.sig_table_dir=config['学员签到表文件夹']
        logo=Image.open(os.path.join(self.public_pic_dir,'01大智小超科学实验室商标.png'))
        self.logo=logo.resize((200*1988//1181,200))
        r,g,b,self.a=self.logo.split()

    def put_mark(self,img_src='e:\\temp\\每周乐高课_学员\\LBC陆炳辰\\20210701-L097会投掷的车-006.JPG'):
        txt=img_src.split('\\')[-1][:-4].split('-')
        crs_name=txt[1][4:]
        crs_date=txt[0][0:4]+'-'+txt[0][4:6]+'-'+txt[0][6:]
        
        font_size=140
        mk_bg_wid=400+(len(crs_name)+1)*font_size
        mk_bg=Image.new('RGBA',(mk_bg_wid,260),'#FFA833')    
        mk_bg.paste(self.logo,(30,30),mask=self.a)
        draw=ImageDraw.Draw(mk_bg)
        draw.text((360,30),'·'+crs_name,fill='#FFFFFF',font=TxtFormat().fonts('华文新魏',font_size))
        date_size=40
        draw.text(((mk_bg_wid-370-TxtFormat().char_len(crs_date))//2+300,190),crs_date,fill='#FFFFFF',font=TxtFormat().fonts('方正韵动粗黑简',date_size))

        mk_bg=pics_modify.circle_corner(mk_bg,radii=150)
        # mk_bg.show()
        r2,g2,b2,a2=mk_bg.split()
        img=Image.open(img_src)
        #缩放至固定的分辨率
        if img.size[0]!=4032:
            img=img.resize((4032,3024))
        img.paste(mk_bg,(80,2600),mask=a2)
        return img

    def pick_pics(self,std_name='LWL廖韦朗',start_date='20210103',end_date='20210506'):
        pic_list=[]
        for fn in os.listdir(os.path.join(self.std_pic_dir,std_name)):
            if fn[-3:].lower()=='jpg' or fn[-4:].lower()=='jpeg':
                date_s = datetime.strptime(start_date, "%Y%m%d") 
                date_e = datetime.strptime(end_date, "%Y%m%d")  
                date_pic=datetime.strptime(fn[:8],"%Y%m%d") 
                if date_pic>=date_s and date_pic<=date_e:
                    pic_list.append(os.path.join(self.std_pic_dir,std_name,fn))

        return pic_list

    def put_simple_marks(self,std_name_list=['LWL廖韦朗'],start_date='20210103',end_date='20210506'):
        print('正在给照片加标签')
        for std_name in std_name_list:
            print('正在处理 {} 的照片……'.format(std_name),end='')
            pic_list=self.pick_pics(std_name,start_date=start_date,end_date=end_date)
            save_dir=os.path.join(self.save_dir,std_name)
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            for pic in pic_list:
                fn=pic.split('\\')[-1][:-4]+'_mark.jpg'
                img=self.put_mark(pic)                
                img.save(os.path.join(save_dir,fn),quality=95,subsampling=0)
            print('完成')
        os.startfile(self.save_dir)
    
    def put(self,term='2021春',weekdays=[1,4],start_date='20210103',end_date='20210506'):
        xlsxs=[]
        print('\n正在生成学员列表……',end='')
        for wd in weekdays:
            wd=days_calculate.num_to_ch(wd)
            fn=term+'-'+'学生信息表（周'+wd+'）.xlsx'
            xls=os.path.join(self.sig_table_dir,self.place_input,'学生信息表',fn)
            xlsxs.append(xls)
        
        std_list=[]
        for xlsx in xlsxs:  
            df=pd.read_excel(xlsx,sheet_name='学生档案表')
            std_name_pre=df['姓名首拼'].str.cat(df['学生姓名'])
            std_name=std_name_pre.to_list()
            std_list.extend(std_name)
        # print(std_list)
        print('完成\n')

        self.put_simple_marks(std_name_list=std_list,start_date=start_date,end_date=end_date)
        print('完成')

class StepMark(pics):
    def __init__(self):
        super().__init__()
        self.para={
            '乐高step1':['认识零件','今天是我第一次上乐高机器人课！老师带我认识了各种认识板、砖、梁、销、齿轮等各种零件。用这些零件可以搭建各种各样的东西，好神奇！'],
            '乐高step2':['搭建环节','我学会了观察对比图纸和按照颜色、结构来选零件和搭建。老师夸奖我观察细致、会独立思考！'],
            '乐高step3':['作品环节','我使用了32个零件，经历20个搭建步骤，过程中会遇到零件扣不稳、操作不慎作品散架等等问题，但我都勇敢面对，完成作品“小汽车”，好开心！'],
            '乐高step4':['分享环节','我勇敢自信地走上前台，跟大家分享我的作品。今天很有成就感，我也是小小工程师啦！'],
            '乐高step5':['归纳整理','老师教了我们按颜色、形状归纳分类的方法，我把小汽车拆卸了，对应零件的格子放好，让零件回到自己的家。老师说学会归纳整理是一个非常好的习惯！']
        }

    
    def read_fn(self):
        decode_tag=pics()
        ptn_std_name=re.compile(r'^[a-zA-Z]+[\u4e00-\u9fa5]+')
        for fn in os.listdir(self.input_dir):
            if fn[-3:].lower()=='jpg' or fn[-4:].lower()=='jpeg':
                img_fn=os.path.join(self.input_dir,fn)
                _tags=iptcinfo3.IPTCInfo(img_fn)
                tags=decode_tag.code_to_str(_tags)
                if len(tags)>0:
                    for tag in tags: 
                        if tag in list(self.para.keys()):
                            txt=self.para[tag]
                            # print(txt)
                            # self.put_step_txt(txt=txt,fn=img_fn)
        return txt
    
    def put_step_txt(self,txt='测试测试',fn='C:\\Users\\jack\\Desktop\\新建文件夹\\20210325-L066弹力小车-030.JPG'):
        img=Image.open(fn)
        draw=ImageDraw.Draw(img)
        draw.text((3000,2500),text=txt, fill = '#2A68B1',font=TxtFormat().fonts('优设标题',100))
        # print(img.size)
        # img.show()
        save_fn=fn.split('\\')[-1]
        # print(save_fn)
        img.save(os.path.join(self.out_dir,save_fn))


    def putCover(self,height=2250,term='2020秋',crop='yes',bigger='yes',weekday=2):
        def read_excel():
            crsFile=['课程信息表.xlsx','课程信息']
            # if weekday==2:
            #     stdFile=['2020乐高课程签到表（周二）.xlsx','学生上课签到表']
            # elif weekday==6:
            #     stdFile=['2020乐高课程签到表（周六）.xlsx','学生上课签到表']
            wd=days_calculate.num_to_ch(weekday)
            stdFile=[term+'-'+'学生信息表（体验）.xlsx','学生上课签到表']
            # stdFile=['2019科学实验课学员档案2.xlsx','学员名单']
            crs=pd.read_excel(os.path.join(self.CrsInfoDir,crsFile[0]),skiprows=0,sheet_name=crsFile[1])
            stds=pd.read_excel(os.path.join(self.StdInfoDir,stdFile[0]),skiprows=1,sheet_name=stdFile[1])
            stds.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名', \
                                'Unnamed: 5':'昵称','Unnamed: 6':'性别','Unnamed: 7':'上期课时结余','Unnamed: 8':'购买课时','Unnamed: 9':'目前剩余课时', \
                                'Unnamed: 10':'上课数量统计汇总'},inplace=True)
            # print(stds)
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
            total_pics_dir=os.path.join(self.totalPics,term,term+'-体验课')
            # if not os.path.exists(total_pics_dir):
            #     os.makedirs(total_pics_dir)

            for fileName in os.listdir(total_pics_dir):
                if fileName[-3:].lower()=='jpg' or fileName[-4:].lower()=='jpeg':
                    fn=fileName.split('-')
                    crsName=fn[1][4:]
                    crsCode=fn[1][0:4]
                    real_addr=os.path.join(total_pics_dir,fileName)
                    tag=self.code_to_str(iptcinfo3.IPTCInfo(real_addr))
                    # print(tag)
                    if len(tag)>0:
                        for _tag in tag:        
                            # print(_tag)     
                            _tag.strip()
                            _tag=_tag.replace(' ','')                            

                            if _tag.lower() in list(self.para.keys()):
                                    txt_step=self.para[_tag.lower()][1]
                                    step_num=_tag.lower()
                                    step_name=self.para[_tag.lower()][0]
                                
                            # print(_tag)
                            if ptn_std_name.match(_tag):
                                # print(_tag)
                                _tag_py=re.findall(r'[a-zA-Z]+',_tag)[0] 
                                _tag_zh=re.findall(r'[\u4e00-\u9fa5]+',_tag)[0] 

                                if _tag_zh in stdList:
                                    # print('80_tag:',_tag)
                                    std_class=stds[stds['学生姓名']==_tag_zh]['班级'].tolist()[0]
                                    std_name=_tag_zh
                                    std_py=_tag_py
                                    knlg=crs[crs['课程编号']==crsCode]['知识点'].tolist()[0]
                        infos.append([real_addr,std_name,crsName,knlg,fileName,std_py,txt_step,std_class,step_num,step_name])
                            
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

            rct=Image.new('RGBA',(w,h),(255,255,255,190))

            draw=ImageDraw.Draw(img)
            img.paste(rct,(0,int(img.size[1]-h)),mask=rct)

            # draw.rectangle([(0,int(img.size[1]-h)),(w,img.size[1])],fill='#eae8e8') #背景
            r=img.size[1]/3024
    
            # print(img.size,w,h)
            title='\n'.join(TxtFormat().split_txt_Chn_eng(ratio(360,r),ratio(90,r),txt[1])[0][0])
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
            # img.paste(pic_qrcode,(ratio(3560,r),ratio(2500,r)))    
            # draw.text((ratio(3560,r),ratio(2880,r)), '微信扫码关注视频号', fill = '#000000',font=TxtFormat().fonts('微软雅黑',ratio(40,r)))

            partTitle=ratio(1000,r)
            partStep=ratio(1800,r)
            titleSize=ratio(300,r)
            knlgSize=ratio(80,r)
            dateSize=ratio(65,r)
            xTitle=ratio(700,r)
            
            # txt[1]='恐龙'
            # print('before  ',titleSize)
            while len(txt[1])*titleSize>partTitle:
                titleSize=titleSize-2
            # print('after  ',titleSize)
            xDate=int(len(txt[1])*titleSize/2+xTitle)-int(TxtFormat().char_len(txt[2])*dateSize/2)-20


            draw.text((xTitle,int(ratio(2530,r)+0.8*(ratio(300,r)-titleSize))), txt[1], fill = '#2A68B1',font=TxtFormat().fonts('优设标题',titleSize))  #课程题目  单个汉字的上方会有空间，空间大小与字体成正比，所以y坐标要根据字体大小改变。
            draw.text((xDate,ratio(2880,r)), txt[2], fill = '#2A68B1',font=TxtFormat().fonts('微软雅黑',dateSize))  #日期，坐标根据课程题目调整，居中对齐
            # draw.text((1800,2500), txt[3], fill = '#2A68B1',font=TxtFormat().fonts('杨任东竹石体',140))  #知识点
            draw.text((ratio(780,r),ratio(2450,r)), '乐高机器人体验课', fill = '#2A68B1',font=TxtFormat().fonts('汉仪字酷堂义山楷w',80))  #知识点

            draw.rectangle((ratio(780,r),ratio(2560,r),ratio(1590,r),ratio(2565,r)),'#2A68B1')


            # txt3Dot=re.sub('\d.','· ',txt[4]) #将数字替换成点
            # print('153:',isinstance(txt3Dot,str))
            TxtFormat().put_txt_img(draw,txt[4],partStep,[ratio(1850,r),ratio(2530,r)],25,'#2A68B1','楷体',knlgSize) #步骤描述


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
                    img=ResizeCrop(info[0],h_min=height,crop=crop,bigger=bigger)
                    if not isinstance(img,str):
                        bg_h,bg_w=int(img.size[1]*0.2018),img.size[0]
                        txt_write=[info[0],info[2],date_crs,info[3],info[6]]
                        # print('173 txt_write:',txt_write)
                        draw(img,bg_w,bg_h,txt_write)
                        # saveDir=os.path.join(self.stdPicsDir,term,'体验课',str(weekday).zfill(2)+info[5]+info[1])
                        saveDir=os.path.join(self.stdPicsDir,term,'体验课打标',info[7]+'_'+info[5]+info[1])
                        save_filename=info[8]+'_'+info[9]+'.jpg'
                        # print(type(save_filename))
                        saveName=os.path.join(saveDir,save_filename)
                        if not os.path.exists(saveDir):
                            # print(saveName)
                            os.makedirs(saveDir)
                            img.save(saveName,quality=95,subsampling=0) #subsampling参数：子采样，通过实现色度信息的分辨率低于亮度信息来对图像进行编码的实践。可能的子采样值是0,1和2。
                        else:
                            img.save(saveName,quality=95,subsampling=0)
                            # print(saveName)
                        # print('生成的照片所在文件夹：{}'.format(saveDir))
                        # print('测试保存:',saveName)
                    else:
                        smallpics.append(info[4])

                if smallpics:
                    msg='完成 {}/{} 个文件。{}个文件太小，未完成：'.format(len(infos)-len(smallpics),len(infos),len(smallpics))+', '.join(smallpics)+'   too small.'
                    print(msg)
                print('完成')
                os.startfile(os.path.join(self.stdPicsDir,term,'体验课打标'))
            else:
                print('无合适的照片')

        putCoverToPics()

class AfterClassMark(pics):
    def __init__(self):
        super().__init__()

    
    def get_pics(self,crs_date='20210924',crs_name='L107我的小房子'):
        pics_addr=[]
        for parent,dirs,fns in os.walk(os.path.join(self.after_class_dir,crs_date+'-'+crs_name)):
            for fn in fns:
                if fn[-3:].lower()=='jpg' or fn[-4:].lower()=='jpeg':
                    pics_addr.append(os.path.join(parent,fn))
        # print(pics_addr)
        return pics_addr

    def resize_crop(self,img,h_min=2250,crop='yes',bigger='yes'):
        # img=Image.open(pic_src)
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

    def ratio(self,p,rate):
            k=int(p*rate)
            return k

    def put_cover_to_pic(self,crs_date='20210924',crs_name='L107我的小房子',img_src='E:\\大智小超\\课后照片及反馈\\20210924-L107我的小房子\\CJY陈锦媛\\20210924-L107我的小房子-017.JPG',forced_ht=''):
        img=Image.open(img_src)
        # img=img.convert('RGBA')
        img=self.resize_crop(img,h_min=2250,crop='yes',bigger='yes')
        w,h=img.size

        #半透明背景
        # rct=Image.new('RGBA',(w,h//10),(255,255,255,190))   
        ft_size_crs=90   
        p_rct_x,p_rct_y= 30,h-h//10
        h_rct=h//13
        paraFormat=TxtFormat()
        w_rct=int(TxtFormat().char_len(crs_name[4:])*ft_size_crs+80)
        rct=Image.new('RGB',(w_rct,h_rct),(255,255,255))
        shape=Shape()
        rct=shape.circle_corner(rct,radii=100)
        rct_blender=Image.new('RGBA',(rct.size[0],rct.size[1]),(0,0,0,0))
        #增加透明度
        rct=Image.blend(rct_blender,rct,0.9) 
        # rct1,rct2,rct3,rct4=rct_alpha.split()
        draw=ImageDraw.Draw(img)
        img.paste(rct,(p_rct_x,p_rct_y),mask=rct)
        # draw.rectangle([(0,int(img.size[1]-h)),(w,img.size[1])],fill='#eae8e8') #背景
        r=h/3024
        # title='\n'.join(TxtFormat().split_txt_Chn_eng(self.ratio(360,r),self.ratio(90,r),txt[1])[0][0])
        _pic_logo=Image.open('I:\\大智小超\\公共素材\\图片类\\logoForPic.png').convert('RGBA')
        _pic_logo2=Image.open('I:\\大智小超\\公共素材\\图片类\\logoDZXC.png').convert('RGBA')
        # pic_logo=_pic_logo.resize((self.ratio(430,r),self.ratio(430/1.7616,r)))
        # pic_logo2=_pic_logo2.resize((self.ratio(350,r),self.ratio(350*46/36,r)))
        # red,g,b,a=pic_logo.split()
        # red2,g2,b2,a2=pic_logo2.split()
        # _pic_qrcode=Image.open('I:\\大智小超\\公共素材\\图片类\\大智小超视频号二维码2.png')
        # print(self.ratio(350,r),self.ratio(350,r))
        # pic_qrcode=_pic_qrcode.resize((self.ratio(350,r),self.ratio(350,r)))
        # img.paste(pic_logo,(self.ratio(50,r),self.ratio(2740,r)),mask=a)
        # img.paste(pic_logo2,(self.ratio(120,r),self.ratio(2500,r)),mask=a2)
        # img.paste(pic_qrcode,(self.ratio(3560,r),self.ratio(2500,r)))    
        # draw.text((self.ratio(3560,r),self.ratio(2880,r)), '微信扫码关注视频号', fill = '#000000',font=TxtFormat().fonts('微软雅黑',self.ratio(40,r)))
        # draw.text((self.ratio(40,r),self.ratio(h-h//10+10,r)), crs_date[:4]+'-'+crs_date[4:6]+'-'+crs_date[6:], fill = '#3d7ab3',font=TxtFormat().fonts('楷体',self.ratio(60,r)))
        p_date_x=p_rct_x+(w_rct-TxtFormat().char_len(crs_date[:4]+'-'+crs_date[4:6]+'-'+crs_date[6:])*50)//2
        draw.text((p_date_x,p_rct_y+h_rct*9//10-40), crs_date[:4]+'-'+crs_date[4:6]+'-'+crs_date[6:], fill = '#3d7ab3',font=TxtFormat().fonts('楷体',50))
        # draw.text((self.ratio(1560,r),self.ratio(2730,r)), crs_name[4:], fill = '#3d7ab3',font=TxtFormat().fonts('优设标题',self.ratio(ft_size_crs,r)))
        p_crs_x=p_rct_x+(w_rct-TxtFormat().char_len(crs_name[4:])*ft_size_crs)//2
        draw.text((p_crs_x,p_rct_y+5), crs_name[4:], fill = '#3d7ab3',font=TxtFormat().fonts('优设标题',ft_size_crs))
        # draw.text((self.ratio(630,r),self.ratio(2730,r)), '科学机器人', fill = '#3db357',font=TxtFormat().fonts('汉仪超级战甲',self.ratio(110,r)))
        # draw.text((self.ratio(630,r),self.ratio(2830,r)), '编程课', fill = '#3db357',font=TxtFormat().fonts('汉仪超级战甲',self.ratio(110,r)))


        # img.show()
        if forced_ht:
            w,h=img.size
            forced_ht=int(forced_ht)
            forced_wid=forced_ht*w//h
            img=img.resize((forced_wid,forced_ht))
        return img
        
    def group_put(self,crs_date='20210924',crs_name='L107我的小房子',forced_ht=''):
        print('正在给照片加上标签……\n')
        pic_list=self.get_pics(crs_date=crs_date,crs_name=crs_name)
        # print(pic_list)
        for pic in tqdm(pic_list):
            img=self.put_cover_to_pic(crs_date=crs_date,crs_name=crs_name,img_src=pic,forced_ht=forced_ht)
            img.save(pic,quality=95,subsampling=0)
        print('完成')

    def putCover(self,height=2250,term='2020秋',crop='yes',bigger='yes',weekday=2):
        def read_excel():
            crsFile=['课程信息表.xlsx','课程信息']
            # if weekday==2:
            #     stdFile=['2020乐高课程签到表（周二）.xlsx','学生上课签到表']
            # elif weekday==6:
            #     stdFile=['2020乐高课程签到表（周六）.xlsx','学生上课签到表']
            wd=days_calculate.num_to_ch(weekday)
            stdFile=[term+'-'+'学生信息表（周'+wd+'）.xlsx','学生上课签到表']
            # stdFile=['2019科学实验课学员档案2.xlsx','学员名单']
            crs=pd.read_excel(os.path.join(self.CrsInfoDir,crsFile[0]),skiprows=0,sheet_name=crsFile[1])
            stds=pd.read_excel(os.path.join(self.StdInfoDir,stdFile[0]),skiprows=1,sheet_name=stdFile[1])
            stds.rename(columns={'Unnamed: 0':'ID','Unnamed: 1':'机构','Unnamed: 2':'班级','Unnamed: 3':'姓名首拼','Unnamed: 4':'学生姓名', \
                                'Unnamed: 5':'昵称','Unnamed: 6':'性别','Unnamed: 7':'上期课时结余','Unnamed: 8':'购买课时','Unnamed: 9':'目前剩余课时', \
                                'Unnamed: 10':'上课数量统计汇总'},inplace=True)
            # print(stds)
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
            total_pics_dir=os.path.join(self.totalPics,term,term+'-每周课程16')
            # if not os.path.exists(total_pics_dir):
            #     os.makedirs(total_pics_dir)

            for fileName in os.listdir(total_pics_dir):
                if fileName[-3:].lower()=='jpg' or fileName[-4:].lower()=='jpeg':
                    fn=fileName.split('-')
                    crsName=fn[1][4:]
                    crsCode=fn[1][0:4]
                    real_addr=os.path.join(total_pics_dir,fileName)
                    tag=self.code_to_str(iptcinfo3.IPTCInfo(real_addr))
                    if len(tag)>0:
                        for _tag in tag:        
                            # print(_tag)     
                            _tag.strip()
                            _tag=_tag.replace(' ','')           
                            if ptn_std_name.match(_tag):
                                _tag_py=re.findall(r'[a-zA-Z]+',_tag)[0] 
                                _tag_zh=re.findall(r'[\u4e00-\u9fa5]+',_tag)[0] 
                                
                            # print('77 _tag:',_tag)

                                if _tag_zh in stdList:
                                    # print('80_tag:',_tag)
                                    std_name=_tag_zh
                                    std_py=_tag_py
                                    knlg=crs[crs['课程编号']==crsCode]['知识点'].tolist()[0]
                                    infos.append([real_addr,std_name,crsName,knlg,fileName,std_py])
            print('完成')
                # print('77infos:',infos)
            return infos

            

        def ratio(p,rate):
            k=int(p*rate)
            return k

        def draw(img,w,h,txt):           

            rct=Image.new('RGBA',(w,h),(255,255,255,190))

            draw=ImageDraw.Draw(img)
            img.paste(rct,(0,int(img.size[1]-h)),mask=rct)

            # draw.rectangle([(0,int(img.size[1]-h)),(w,img.size[1])],fill='#eae8e8') #背景
            r=img.size[1]/3024
    
            # print(img.size,w,h)
            title='\n'.join(TxtFormat().split_txt_Chn_eng(ratio(360,r),ratio(90,r),txt[1])[0][0])
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
            draw.text((ratio(3560,r),ratio(2880,r)), '微信扫码关注视频号', fill = '#000000',font=TxtFormat().fonts('微软雅黑',ratio(40,r)))

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
            xDate=int(len(txt[1])*titleSize/2+xTitle)-int(TxtFormat().char_len(txt[2])*dateSize/2)-20


            draw.text((xTitle,int(ratio(2400,r)+0.8*(ratio(300,r)-titleSize))), txt[1], fill = '#2A68B1',font=TxtFormat().fonts('优设标题',titleSize))  #课程题目  单个汉字的上方会有空间，空间大小与字体成正比，所以y坐标要根据字体大小改变。
            draw.text((xDate,ratio(2800,r)), txt[2], fill = '#2A68B1',font=TxtFormat().fonts('微软雅黑',dateSize))  #日期，坐标根据课程题目调整，居中对齐
            # draw.text((1800,2500), txt[3], fill = '#2A68B1',font=TxtFormat().fonts('杨任东竹石体',140))  #知识点
            draw.text((ratio(1900,r),ratio(2450,r)), '课程知识点', fill = '#2A68B1',font=TxtFormat().fonts('汉仪字酷堂义山楷w',80))  #知识点

            draw.rectangle((ratio(1900,r),ratio(2560,r),ratio(2400,r),ratio(2565,r)),'#2A68B1')


            txt3Dot=re.sub('\d.','· ',txt[3]) #将数字替换成点
            # print('153:',isinstance(txt3Dot,str))
            TxtFormat().put_txt_img(draw,txt3Dot,partKnlg,[ratio(1850,r),ratio(2620,r)],25,'#2A68B1','楷体',knlgSize) #知识点，可设置行间距

        def putCoverToPics():
            infos=self.get_pics()
            # print(infos)
            # infos=[infos[2]]            
            if infos:
                print('正在写入标注信息并按姓名保存到文件夹')
                smallpics=[]
                for info in tqdm(infos):
                    date_crs=info[4].split('-')[0][0:4]+'-'+info[4].split('-')[0][4:6]+'-'+info[4].split('-')[0][6:]
                    # img=Image.open(info[0])
                    img=ResizeCrop(info[0],h_min=height,crop=crop,bigger=bigger)
                    if not isinstance(img,str):
                        bg_h,bg_w=int(img.size[1]*0.2018),img.size[0]
                        txt_write=[info[0],info[2],date_crs,info[3]]
                        # print('173 txt_write:',txt_write)
                        draw(img,bg_w,bg_h,txt_write,mode=mode,txt2=txt2)
                        saveDir=os.path.join(self.stdPicsDir,term,'冲印版',str(weekday).zfill(2)+info[5]+info[1])
                        saveName=os.path.join(saveDir,info[4])
                        if not os.path.exists(saveDir):
                            # print(saveName)
                            os.makedirs(saveDir)
                            img.save(saveName,quality=95,subsampling=0) #subsampling参数：子采样，通过实现色度信息的分辨率低于亮度信息来对图像进行编码的实践。可能的子采样值是0,1和2。
                        else:
                            img.save(saveName,quality=95,subsampling=0)
                            # print(saveName)
                        # print('生成的照片所在文件夹：{}'.format(saveDir))
                        # print('测试保存:',saveName)
                    else:
                        smallpics.append(info[4])

                if smallpics:
                    msg='完成 {}/{} 个文件。{}个文件太小，未完成：'.format(len(infos)-len(smallpics),len(infos),len(smallpics))+', '.join(smallpics)+'   too small.'
                    print(msg)
                print('完成')
            else:
                print('无合适的照片')

        putCoverToPics()

class TempMark:
    def __init__(self):
        print('正在初始化参数……',end='')
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'LegoStudentPic.config'),'r',encoding='utf-8') as f:
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
        self.after_class_dir=config['课后照片及反馈文件夹']
        # print(self.publicPicDir,self.CrsInfo)
        print('完成')
        

    def putCover(self,input_dir='C:\\Users\\jack\\Desktop\\7寸相片冲印',height=2250,crop='yes',bigger='yes'):
        def read_excel():
            crsFile=['课程信息表.xlsx','课程信息']
            # if weekday==2:
            #     stdFile=['2020乐高课程签到表（周二）.xlsx','学生上课签到表']
            # elif weekday==6:
            #     stdFile=['2020乐高课程签到表（周六）.xlsx','学生上课签到表']

    
            crs=pd.read_excel(os.path.join(self.CrsInfoDir,crsFile[0]),skiprows=0,sheet_name=crsFile[1])
     
            # print(stds)
            # std=stds[stds['学生姓名']==stdName]
            # std_basic=std[['姓名首拼','学生姓名']]
            # std_crs=std[std.iloc[:,:]=='√'].dropna(axis=1)
            # std_res=pd.concat([std_basic,std_crs],axis=1)
            # print(crs)
            return crs

        def read_pics_new():
            print('正在读取照片……',end='')
            crs=read_excel()

            infos=[]
            for fileName in os.listdir(input_dir):
                # print(fileName)
                if fileName[-3:].lower()=='jpg' or fileName[-4:].lower()=='jpeg':
                    fn=fileName.split('-')
                    crsName=fn[1][4:]
                    crsCode=fn[1][0:4]
                    real_addr=os.path.join(input_dir,fileName)
                    knlg=crs[crs['课程编号']==crsCode]['知识点'].tolist()[0]
                    infos.append([real_addr,crsName,knlg,fileName])
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

            rct=Image.new('RGBA',(w,h),(255,255,255,190))

            draw=ImageDraw.Draw(img)
            img.paste(rct,(0,int(img.size[1]-h)),mask=rct)

            # draw.rectangle([(0,int(img.size[1]-h)),(w,img.size[1])],fill='#eae8e8') #背景
            r=img.size[1]/3024
    
            # print(img.size,w,h)
            # print(txt)
            # title='\n'.join(TxtFormat().split_txt_Chn_eng(ratio(360,r),ratio(90,r),txt[1])[0][0])
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
            draw.text((ratio(3560,r),ratio(2880,r)), '微信扫码关注视频号', fill = '#000000',font=TxtFormat().fonts('微软雅黑',ratio(40,r)))

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
            xDate=int(len(txt[1])*titleSize/2+xTitle)-int(TxtFormat().char_len(txt[2])*dateSize/2)-20


            draw.text((xTitle,int(ratio(2400,r)+0.8*(ratio(300,r)-titleSize))), txt[1], fill = '#2A68B1',font=TxtFormat().fonts('优设标题',titleSize))  #课程题目  单个汉字的上方会有空间，空间大小与字体成正比，所以y坐标要根据字体大小改变。
            draw.text((xDate,ratio(2800,r)), txt[2], fill = '#2A68B1',font=TxtFormat().fonts('微软雅黑',dateSize))  #日期，坐标根据课程题目调整，居中对齐
            # draw.text((1800,2500), txt[3], fill = '#2A68B1',font=TxtFormat().fonts('杨任东竹石体',140))  #知识点
            draw.text((ratio(1900,r),ratio(2450,r)), '课程知识点', fill = '#2A68B1',font=TxtFormat().fonts('汉仪字酷堂义山楷w',80))  #知识点

            draw.rectangle((ratio(1900,r),ratio(2560,r),ratio(2400,r),ratio(2565,r)),'#2A68B1')


            txt3Dot=re.sub('\d.','· ',txt[3]) #将数字替换成点
            # print('153:',isinstance(txt3Dot,str))
            TxtFormat().put_txt_img(draw,txt3Dot,partKnlg,[ratio(1850,r),ratio(2620,r)],25,'#2A68B1','楷体',knlgSize) #知识点，可设置行间距

        def putCoverToPics():
            infos=read_pics_new()
            #real_addr,crsName,knlg,fileName
            # print(infos)
            # infos=[infos[2]]            
            if infos:
                print('正在写入标注信息并按姓名保存到文件夹')
                smallpics=[]
                for info in tqdm(infos):
                    date_crs=info[3].split('-')[0][0:4]+'-'+info[3].split('-')[0][4:6]+'-'+info[3].split('-')[0][6:]
                    # img=Image.open(info[0])
                    img=ResizeCrop(info[0],h_min=height,crop=crop,bigger=bigger)
                    if not isinstance(img,str):
                        bg_h,bg_w=int(img.size[1]*0.2018),img.size[0]
                        txt_write=[info[0],info[1],date_crs,info[2]]
                        # print('173 txt_write:',txt_write)
                        draw(img,bg_w,bg_h,txt_write)
                        saveDir=os.path.join(self.stdPicsDir,'临时打标')
                        saveName=os.path.join(saveDir,info[3])
                        if not os.path.exists(saveDir):
                            # print(saveName)
                            os.makedirs(saveDir)
                            img.save(saveName,quality=95,subsampling=0) #subsampling参数：子采样，通过实现色度信息的分辨率低于亮度信息来对图像进行编码的实践。可能的子采样值是0,1和2。
                        else:
                            img.save(saveName,quality=95,subsampling=0)
                            # print(saveName)
                        # print('生成的照片所在文件夹：{}'.format(saveDir))
                        # print('测试保存:',saveName)
                    else:
                        smallpics.append(info[3])

                if smallpics:
                    msg='完成 {}/{} 个文件。{}个文件太小，未完成：'.format(len(infos)-len(smallpics),len(infos),len(smallpics))+', '.join(smallpics)+'   too small.'
                    print(msg)
                print('完成')
            else:
                print('无合适的照片')
            
            os.startfile(saveDir)
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
    pic.putCover(height=2250,term='2022春',crop='yes',bigger='yes',weekday=5,savemode='all_together') #savemode参数：all_together将照片放同一个文件夹，individual将照片按学生姓名分开放
    # stu_pics=LegoStudentPicDistribute.LegoPics(crsDate=20210909,crsName='L106手动小赛车',weekday=4,term='2021秋',mode='tiyan')
    # stu_pics.dispatch()

    #体验课加步骤
    # p=StepMark()
    # p.putCover(height=2250,term='2021秋',crop='yes',bigger='yes',weekday=5)

    #课后照片加课程名称及时间
    # p=AfterClassMark()
    # p.get_pics(crs_date='20210924',crs_name='L107我的小房子')
    # p.put_cover_to_pic(img_src='e:/temp/20210924-L107我的小房子-080.JPG')
    # p.group_put(crs_date='20210924',crs_name='L107我的小房子')

    # p=TempMark()
    # p.putCover(input_dir='C:\\Users\\jack\\Desktop\\7寸相片冲印',height=2250,crop='yes',bigger='yes')

    # pic=SimpleMark(place_input='001-超智幼儿园')
    # # pic.put_mark()
    # # pic.put_simple_marks(std_name_list=['LWL廖韦朗','LBC陆炳辰'],start_date='20210103',end_date='20210506')
    # pic.put(term='2021春',weekdays=[1,4],start_date='20210103',end_date='20210506')