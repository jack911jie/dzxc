import os
import sys
sys.path.append('i:/py/dzxc/module')
import readConfig
import re
import json
import moviepy.editor as mpy
import moviepy.audio.fx.all as afx
import pandas as pd
import numpy as np
from PIL import Image
from PIL import ImageDraw  
from PIL import ImageFont
import psutil



class LegoCons:
    def __init__(self,pth,consName):
        self.pth=pth
        self.consName=consName
        self.creatItem()
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'picToMP4config.txt'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
            config=json.loads(_line)
        self.crsInfo=config['课程信息表']
        
    def creatItem(self):
        New=os.path.join(self.pth,self.consName)
        if not os.path.exists(New):
            print('新项目，创建文件夹...\n')
            os.makedirs(New)
            new_totals=os.path.join(self.pth,self.consName,'零件总图')
            new_video=os.path.join(self.pth,self.consName,'video')
            print('文件夹{}创建完成\n'.format(' '+self.consName+' '))
            
            if not os.path.exists(new_totals):
                os.makedirs(new_totals)
                print('文件夹{}创建完成\n'.format(' 零件总图 '))
            if not os.path.exists(new_video):
                os.makedirs(new_video)
                print('文件夹{}创建完成\n'.format(' video '))
        else:
            print('已有项目\n')
            new_totals=os.path.join(self.pth,self.consName,'零件总图')
            new_video=os.path.join(self.pth,self.consName,'video')
            if not os.path.exists(new_totals):
                os.makedirs(new_totals)
                print('文件夹{}创建完成\n'.format(' 零件总图 '))
            if not os.path.exists(new_video):
                os.makedirs(new_video)
                print('文件夹{}创建完成\n'.format(' video '))
          
    def expHTML_lego(self):
        lsts=self.makeList()     
        
        html_pic_addr=''
        _pre_pic='''
        </section>
                        <section style="color: #bfbfbf;padding-top: 10px;padding-bottom: 10px;display: inline-block;width: 100%;box-sizing: border-box;" data-width="75%">
                            <section class="_135editor" data-tools="135编辑器" data-id="95138">
                                <section class="_135editor">
                                    <section style="margin: 1em auto;text-align: center;padding: 5px;border-width: 1px;border-style: solid;border-color: transparent;overflow: hidden;box-sizing: border-box;">
                                        <section class="135brush" data-style="display: inline-block;width: 100%;margin:0;padding:0;" style="white-space: nowrap;overflow-x: scroll;">
        '''


        pre_pic_addr='''
         <img class="" data-ratio="0.75" data-type="jpeg" data-w="1200" data-width="100%" src="https://chuntianhuahua-1257410889.cos.ap-guangzhou.myqcloud.com/legoCons/
         '''
        after_pic_addr='"/>'
        
        for lst in lsts:
            html_pic_addr=html_pic_addr+pre_pic_addr.strip()+lst+after_pic_addr+'\n'
            
        out=_pre_pic+html_pic_addr
        print('生成html...完成\n')
        
        with open(os.path.join(self.pth,self.consName,self.consName+'_html.txt'),'w') as f:
            f.write(out)
        
        print('写入txt...完成\n')
        
    def putTag(self):        
        dirs=os.path.join(self.pth,self.consName,'零件总图')
        df=pd.read_excel(os.path.join(self.pth,self.consName,self.consName+'步骤列表.xlsx'))
        tag_list=np.array(df.loc[~pd.isna(df['描述'])]).tolist()
        
        print('正在给图片添加文字描述……\n')

        if self.consName[0].upper()=='L': #wedo文字颜色
            txtColor='#6AB34A'
        elif self.consName[0].upper()=='N': #9686文字颜色
            txtColor='#06419a'
        else:
            txtColor='#06419a'
        
        font = ImageFont.truetype('C:\Windows\Fonts\msyh.ttc',80)
        
        for lst in tag_list:
            # if not os.path.exists(os.path.join(self.pth,lst[0])):
            #     pth=os.path.join(self.pth,lst[0].replace('_tagged.png','.png')) #因makelist()生成的列表为包含了tagged.png的文件名，所以要改回来。
            # else:
            #     pth=os.path.join(self.pth,lst[0])
            abs_picPath=os.path.join(self.pth,lst[0].replace('_tagged.png','.png'))
            txt=lst[1]
            # print(abs_picPath,txt)
            img=Image.open(abs_picPath)
            
            draw = ImageDraw.Draw(img)            
            draw.text((200,int(img.size[1]*0.87)), txt, fill = txtColor,font=font)  #6AB34A  95ff67
            # NewName=pth.replace('.png','_tagged.png')
            ptn='.*x.png'
            if re.findall(ptn,abs_picPath):
                NewName=abs_picPath[:-4]+'_tagged.png'
            img.save(NewName)
            #             img.show()       
            #             break
        print('添加文字完成 \n')
    
    def putTag_wedo_coverpage(self):    #已不需要这步，第一张片头的字幕在moviepy中添加，不需另外生成一张图片。
        print('正在为封面首图添加知识点……',end='')
        font_title = ImageFont.truetype('j:/fonts/yousheTitleHei.ttf',80)
        font_title2 = ImageFont.truetype('j:/fonts/yousheTitleHei.ttf',30)
        font_info = ImageFont.truetype('j:/fonts/yousheTitleHei.ttf',25)
        # course_file=self.crsInfo
        df=pd.read_excel(self.crsInfo)
        course_info=df[df['课程编号']==self.consName[0:4]].values.tolist()
        crs_bigtype,crs_type,crs_code,crs_name,crs_age,crs_star,crs_intro,crs_comments,crs_lego,crs_preperation=course_info[0]
        crs_age='适合年龄： '+crs_age
        crs_lego='使用教具： '+crs_lego
        img=Image.open(os.path.join(self.pth,self.consName,self.consName[4:]+'.jpg'))
        draw = ImageDraw.Draw(img)     
        draw.text((100,int(img.size[1]*0.1)), crs_name, fill = '#6AB34A',font=font_title)  #6AB34A  95ff67
        draw.text((100,int(img.size[1]*0.3)), crs_age, fill = '#6AB34A',font=font_title2)  #6AB34A  95ff67
        draw.text((100,int(img.size[1]*0.4)), crs_lego, fill = '#6AB34A',font=font_title2)  #6AB34A  95ff67
        draw.text((100,int(img.size[1]*0.5)), crs_intro, fill = '#6AB34A',font=font_info)  #6AB34A  95ff67
        img.save(os.path.join(self.pth,self.consName,self.consName[4:]+'_cover.jpg'))
        print('完成 \n')        
      
    def makeList(self):
        n=0
        pngList=[]
        for file in os.listdir(os.path.join(self.pth,self.consName)):
            if file[-3:].lower()=='png' and len(file)<10:
                os.rename(os.path.join(self.pth,self.consName,file),os.path.join(self.pth,self.consName,file.zfill(10)))
                pngList.append(file.zfill(10))
                n+=1
            elif file[-3:].lower()=='png':
                pngList.append(self.consName+"/"+file)
        pngList.sort()
        if n>0:
            print('改名{}个文件'.format(n))
        
        if os.path.exists(os.path.join(self.pth,self.consName,'零件总图')):
            #print('存在零件总图\n')
            totalList=[]
            n=0
            for file in os.listdir(os.path.join(self.pth,self.consName,'零件总图')):
                if file[-10:].lower()=='tagged.png': #把加标签的放入列表
                    totalList.append(self.consName+'/零件总图/'+file)
                    n+=1
            if n==0:  #如果没有打过标签的零件图，则把原图放到列表
                for file in os.listdir(os.path.join(self.pth,self.consName,'零件总图')):
                    if file[-3:].lower()=='png': #如无加过标签，列表中也要改为_tagged.png的后缀
                        newFile=file.replace('.png','_tagged.png')
                        totalList.append(self.consName+'/零件总图/'+newFile)
        else:
            #pass
            print('无零件总图')
        totalList.sort()
        
        outList=[self.consName+'/'+self.consName[4:]+'.jpg']
        outList.extend(totalList)
        outList.extend(pngList)

        stepList=os.path.join(self.pth,self.consName,self.consName+'步骤列表.xlsx')
        if not os.path.exists(stepList):
            df=pd.DataFrame(outList)
            df.columns=['步骤']
            df['描述']=''
            df.loc[df['步骤'].str.contains('零件总图'),'描述']='结构分解图 / 零件清单' #将零件总图的“描述”先预设为“零件清单”
            df.to_excel(stepList,index=False)       

        print('生成 文件列表...完成\n')
        return outList    
    
class ConsMovie:
    def __init__(self,pth,consName):     
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'picToMP4config.txt'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
            config=json.loads(_line)

        self.crs_list=config['课程信息表']   
        self.endV=r'I:\大智小超\公共素材\视频类\片尾01.mp4'
        self.pth=pth
        self.consName=consName
        self.k=1/3 #转场占时长的比例

        if consName[0].upper()=='N' or consName[0].upper()=='L':          
            self.bgm=r'I:\大智小超\公共素材\声音类\lego_wedo_long.mp3'
        else:  
            self.bgm=r'I:\大智小超\公共素材\声音类\legoBGM.mp3'    
        
          
        df=pd.read_excel(self.crs_list)
        course_info=df[df['课程编号']==self.consName[0:4]].values.tolist()
        # print(course_info[0])
        self.crs_bigtype,self.crs_type,self.crs_code,self.crs_name,self.crs_age,self.crs_star,self.crs_intro,self.comment,self.crs_lego,self.preperation=course_info[0]
        self.crs_age="适合年龄："+self.crs_age
        self.crs_star="搭建难度："+self.crs_star
        self.crs_lego="使用教具："+self.crs_lego
        
        a=LegoCons(self.pth,self.consName)
        self.lst=a.makeList()
        
    def exportMovie(self):
        if 3*(1-self.k)*len(self.lst) < 56:
            print('60秒以内\n')
            self.expMovie()
        else:
            print('超过60秒，将加快到60秒以内\n')
            self.expMovieIn60s()
        
    
    def expMovie(self):
        print('正在剪辑...\n')
        w,h=850,480
        self.w=w
        self.h=h
        clips=[]
        n=0
        drtn=3
        self.drtn=drtn
        crstime=1
        for p in self.lst:
            fn=os.path.join(self.pth,p)
            if n==0:
                _img=mpy.ImageClip(fn).set_fps(25).set_duration(drtn).resize((w,h))
                clips.append(_img)
                n+=1
            else:
                _img=mpy.ImageClip(fn).set_fps(25).set_duration(drtn).resize((w,h)).crossfadein(crstime).set_start((drtn-crstime)*n)
                clips.append(_img)
                n+=1    
                
        #计算正片时间，供logo字幕用
        drt=0
        for c in clips:
            drt=drt+(c.duration-crstime)
        subtitle_drt=drt-clips[0].duration                            
                
        #加上片尾
        print('正在处理片尾...\n')
        endvideo=mpy.VideoFileClip(self.endV,target_resolution=(h,w)).set_start((drtn-crstime)*n) # 片尾   分辨率是先写h,再写w  大坑
        clips.append(endvideo)        
        
        print('正在处理字幕...\n')        
        clips=self.put_cover_text(clips)        
        
        #text=self.subtitle_right_btm(self.consName[3:],60,w,h)
        #clips.append(text)
    #         text = mpy.TextClip(txt='大智小超科学实验室', fontsize=15, \
    #                  font='Microsoft-YaHei-&-Microsoft-YaHei-UI',color='#9EACC1') \
    #                 .set_pos((685,445)).set_start(clips[0].duration).set_duration(subtitle_drt)        
    #         clips.append(text)      
                
        # finalclip = concatenate_videoclips([_img1,_img2])
        
        logo=mpy.ImageClip(r"I:\大智小超\公共素材\图片类\00大智小超科学实验室商标.png") \
            .set_fps(25).set_duration(drt).resize((80,48)).set_position((740,420)) \
            .crossfadein(crstime).set_start(0)
        clips.append(logo)
        
        print('正在拼接视频...\n') 
        
        finalclip = mpy.CompositeVideoClip(clips)
        
        #加上背景音乐
        print('正在添加背景音乐...\n')
        BGM=mpy.AudioFileClip(self.bgm)
        print([finalclip.duration,endvideo.duration])
        
        final_audio = mpy.CompositeAudioClip([BGM,finalclip.audio]).set_duration(finalclip.duration-endvideo.duration).fx(afx.audio_fadeout,0.8)
        mix=finalclip.set_audio(final_audio)
        totalTime=finalclip.duration
        
        out=os.path.join(self.pth,self.consName,self.consName+'_搭建视频_'+str(int(totalTime))+'s.mp4')
        print('正在导出视频：{}...\n'.format(out))
        mix.write_videofile(out)
        self.killProcess()
        print('Done')
        
    def expMovieIn60s(self):
        print('正在剪辑...\n')
        w,h=850,480
        self.w=w
        self.h=h
        clips=[]
        n=0
        
        drtn=56/len(self.lst)/(1-self.k)
        self.drtn=drtn
        crstime=drtn*self.k
        for p in self.lst:
            fn=os.path.join(self.pth,p)
            if n==0:
                _img=mpy.ImageClip(fn).set_fps(25).set_duration(drtn).resize((w,h))
                clips.append(_img)
                n+=1
            else:
                _img=mpy.ImageClip(fn).set_fps(25).set_duration(drtn).resize((w,h)).crossfadein(crstime).set_start((drtn-crstime)*n)
                clips.append(_img)
                n+=1    
                
        #计算正片时间，供logo字幕用
        drt=0
        for c in clips:
            drt=drt+(c.duration-crstime)
        subtitle_drt=drt-clips[0].duration                            
                
        #加上片尾
        print('正在处理片尾...\n')
        endvideo=mpy.VideoFileClip(self.endV,target_resolution=(h,w)).set_start((drtn-crstime)*n) # 片尾   分辨率是先写h,再写w  大坑
        clips.append(endvideo)        
        
        print('正在处理字幕...\n')
        
        clips=self.put_cover_text(clips) 
        
        logo=mpy.ImageClip(r"I:\大智小超\公共素材\图片类\00大智小超科学实验室商标.png") \
            .set_fps(25).set_duration(drt).resize((80,48)).set_position((740,420)) \
            .crossfadein(crstime).set_start(0)
        clips.append(logo)
        
        print('正在拼接视频...\n')
        finalclip = mpy.CompositeVideoClip(clips)
        
        #加上背景音乐
        print('正在添加背景音乐...\n')
        BGM=mpy.AudioFileClip(self.bgm).set_duration(finalclip.duration-endvideo.duration).fx(afx.audio_fadeout,0.8)
        final_audio = mpy.CompositeAudioClip([BGM,finalclip.audio])
        mix=finalclip.set_audio(final_audio)
        totalTime=finalclip.duration
        
        out=os.path.join(self.pth,self.consName,self.consName+'_搭建视频_forced_60s.mp4')
        print('正在导出视频：{}...\n'.format(out))
        mix.write_videofile(out)
        self.killProcess()
        print('Done')
        
    # 杀死moviepy产生的特定进程
    def killProcess(self):
        # 处理python程序在运行中出现的异常和错误
        try:
            # pids方法查看系统全部进程
            pids = psutil.pids()
            for pid in pids:
                # Process方法查看单个进程
                p = psutil.Process(pid)
                # print('pid-%s,pname-%s' % (pid, p.name()))
                # 进程名
                if p.name() == 'ffmpeg-win64-v4.1.exe':
                    # 关闭任务 /f是强制执行，/im对应程序名
                    cmd = 'taskkill /f /im ffmpeg-win64-v4.1.exe  2>nul 1>null'
                    # python调用Shell脚本执行cmd命令
                    os.system(cmd)
        except:
            pass
        
    def subtitle_right_btm(self,txt,fsz,w,h,drtn):
        x=w-len(txt)*fsz-30
        y=h-fsz-30
        text = mpy.TextClip(txt=txt, fontsize=fsz, font='Microsoft-YaHei-&-Microsoft-YaHei-UI',color='#95ff67').set_pos((x,y)).set_duration(drtn)
        return text
    
    def put_cover_text(self,clips):
        if self.consName[0].upper()=='L':  #wedo课程
            crs_h_name=0.05
            crs_h_age=0.3
            crs_h_lego=0.35
            crs_h_intro=0.6
            clr='#6AB34A'
        elif self.consName[0].upper()=='N': #9686课程
            crs_h_name=0.05
            crs_h_age=0.3
            crs_h_lego=0.35
            crs_h_intro=0.6
            clr='#06419A'
        else:
            crs_h_name=0.05
            crs_h_age=0.6
            crs_h_lego=0.65
            crs_h_intro=0.8
            clr='#6AB34A'
        
        text = mpy.TextClip(txt=self.crs_name, fontsize=85, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_pos((self.w*0.1,self.h*crs_h_name)).set_duration(self.drtn)
        clips.append(text)
        
        text = mpy.TextClip(txt=self.crs_age, fontsize=20, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_pos((self.w*0.1,self.h*crs_h_age)).set_duration(self.drtn)
        clips.append(text)
        
        text = mpy.TextClip(txt=self.crs_lego, fontsize=20, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_pos((self.w*0.1,self.h*crs_h_lego)).set_duration(self.drtn)
        clips.append(text)
        
        text = mpy.TextClip(txt=self.crs_intro,align='West',fontsize=25, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_pos((self.w*0.1,self.h*crs_h_intro)).set_duration(self.drtn)
        clips.append(text)
        
        return clips       
    
class LegoWeekly:
    def __init__(self,pth):
        self.pth=pth
        self.bg=os.path.join(pth,'legoWeeklyBG.jpg')
        
    def expPoster(self):
        print('开始生成积木搭建海报 \n')
        xls=os.path.join(self.pth,'topic.xlsx')
        df=pd.read_excel(xls)
        df_list=list(df['文字/图片地址'])
        bg=Image.open(self.bg)
        topicImg=Image.open(df_list[1])
        
        if df_list[2][-3:].lower()=='png' or df_list[2][-3:].lower()=='jpg':
            print('图片格式的二维码，直接嵌入背景。\n')
            qrImg=Image.open(df_list[2])
            qrImg=qrImg.resize((280,280),Image.ANTIALIAS)
        else:            
            qrImg=self.makeQR(df_list[2])
        
        bg.paste(topicImg,(260,580))
        bg.paste(qrImg,(400,1200))       
        
        
        #行一：主题名字
        font_size=90
        txt=df_list[0]
        lth=self.cal_len_str(txt)
        x=(1080-int((font_size*lth[0]+int(font_size*lth[1])/2)))/2
        font = ImageFont.truetype('C:\Windows\Fonts\msyh.ttc',font_size)
        draw = ImageDraw.Draw(bg)         
        draw.text((x,390), txt, fill = '#009b5b',font=font)
        
        #测试用
    #         font_size=90
    #         txt='海aa龟'
    #         lth=self.cal_len_str(txt)
    #         x=(1080-int((font_size*lth[0]+int(font_size*lth[1])/2)))/2
    #         print(x)
    #         font = ImageFont.truetype('C:\Windows\Fonts\msyh.ttc',font_size)
    #         draw = ImageDraw.Draw(bg)         
    #         draw.text((x,490), txt, fill = '#009b5b',font=font)
        
        #行二
        font_size=40
        txt=df_list[3]
        lth=self.cal_len_str(txt)
        x=(1080-int((font_size*lth[0]+int(font_size*lth[1])/2)))/2
        font = ImageFont.truetype('C:\Windows\Fonts\msyh.ttc',font_size)
        draw = ImageDraw.Draw(bg)         
        draw.text((x,1560), txt, fill = '#003e69',font=font)
        
        #行三
        font_size=65
        txt=df_list[4]
        lth=self.cal_len_str(txt)
        x=(1080-int((font_size*lth[0]+int(font_size*lth[1])/2)))/2
        font = ImageFont.truetype('C:\Windows\Fonts\msyh.ttc',font_size)
        draw = ImageDraw.Draw(bg)         
        draw.text((x,1620), txt, fill = '#003e69',font=font)
        
        #行四
        font_size=35
        txt=df_list[5]
        lth=self.cal_len_str(txt)
        x=(1080-int((font_size*lth[0]+int(font_size*lth[1])/2)))/2
        font = ImageFont.truetype('C:\Windows\Fonts\msyh.ttc',font_size)
        draw = ImageDraw.Draw(bg)         
        draw.text((x,1720), txt, fill = '#003e69',font=font)
        
        bg.save(os.path.join(self.pth,'exportLegoWeek.jpg'))
        print('海报已完成 \n')
        bg.show()
        
        
    def cal_len_str(self,t):
        B=int((len(t.encode('utf-8'))-len(t))/2)
        b=len(t)-B
    #         num_spc=len(re.findall(' ',t))
    #         b=b-num_spc+num_spc*2        
    #         print(B,b)
        return([B,b])

    def makeQR(self,url):
        print('正在转换网址为二维码',end='')
        # 初步生成二维码图像
        qr = qrcode.QRCode(version=5,error_correction=qrcode.constants.ERROR_CORRECT_H,box_size=8,border=4)
        qr.add_data(url)
        qr.make(fit=True)

        # 获得Image实例并把颜色模式转换为RGBA
        img = qr.make_image(fill_color="#003e69")
        img = img.convert("RGBA")

        # 打开logo文件
        icon = Image.open("I:\大智小超\公共素材\图片类\大智小超logo.png")

        # 计算logo的尺寸
        img_w,img_h = img.size
        factor = 5
        size_w = int(img_w / factor)
        size_h = int(img_h / factor)

        # 比较并重新设置logo文件的尺寸
        icon_w,icon_h = icon.size
        if icon_w >size_w:
            icon_w = size_w
        if icon_h > size_h:
            icon_h = size_h
        icon = icon.resize((icon_w,icon_h),Image.ANTIALIAS)

        # 计算logo的位置，并复制到二维码图像中
        w = int((img_w - icon_w)/2)
        h = int((img_h - icon_h)/2)
        icon = icon.convert("RGBA")
        img.paste(icon,(w,h),icon)
        img=img.resize((280,280))

        # 保存二维码
        img.save('./createlogo.png')
        print('……完成\n')
        return img

class BuildAnimation:
    def __init__(self,crs_name,save_yn='yes'):
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)),'picToMP4config.txt'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
            config=json.loads(_line)

        if crs_name[0].upper()=='N' or crs_name[0].upper()=='L':          
            self.bgm=r'I:\大智小超\公共素材\声音类\lego_wedo_long.mp3'
        else:  
            self.bgm=r'I:\大智小超\公共素材\声音类\legoBGM.mp3'  

        self.endV=r'I:\大智小超\公共素材\视频类\片尾01.mp4'
        self.crs_list=config['课程信息表']
        self.src_dir='I:\\乐高\\图纸'
        self.crs_name=crs_name
        self.pics_dir=os.path.join(self.src_dir,self.crs_name,'animation')
        self.crs_code=crs_name[0:4]
        self.w,self.h=850,480
        self.drtn=2
        self.save_yn=save_yn

    def read_pics(self):
        pics_list=[]

        for fn in os.listdir(self.pics_dir):
            pics_list.append(os.path.join(self.pics_dir,fn))

        # print(pics_list)
        # test=pics_list[20:25]
        return pics_list

    def read_excel(self):
        crs_code=self.crs_code
        df=pd.read_excel(self.crs_list)
        crs=df.loc[df['课程编号']==crs_code]   

        knowledge=list(crs['知识点'])
        script=list(crs['课程描述'])
        age=list(crs['年龄'])[0]
        dif_level=list(crs['难度'])
        instrument=list(crs['教具'])
        crs_info=[self.crs_name[4:],age,knowledge[0],script[0],dif_level[0],instrument[0]]      
        stars=crs_info[-1].replace('*','★')
        crs_info[-1]=stars 
        return crs_info


    def put_cover_text(self,clips):
        print('正在添加字幕……',end='')
        if self.crs_name[0].upper()=='L':  #wedo课程
            crs_h_name=0.05
            crs_h_age=0.3
            crs_h_lego=0.4
            crs_h_intro=0.6
            clr='#6AB34A'
        elif self.crs_name[0].upper()=='N': #9686课程
            crs_h_name=0.05
            crs_h_age=0.3
            crs_h_lego=0.35
            crs_h_intro=0.6
            clr='#06419A'
        else:
            crs_h_name=0.05
            crs_h_age=0.6
            crs_h_lego=0.65
            crs_h_intro=0.8
            clr='#6AB34A'
        
        crs_info=self.read_excel()
        crs_name,age,knowledge,script,dif_level,instrument=crs_info

        text = mpy.TextClip(txt=crs_name, fontsize=70, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_fps(25).set_position((self.w*0.1,self.h*crs_h_name)).set_duration(self.drtn).set_start(0)
        clips.append(text)
        
        text = mpy.TextClip(txt='适合年龄：'+age , fontsize=20, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_fps(25).set_position((self.w*0.1,self.h*crs_h_age)).set_duration(self.drtn).set_start(0)
        clips.append(text)
        
        text = mpy.TextClip(txt='使用教具：'+instrument, fontsize=20, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_fps(25).set_position((self.w*0.1,self.h*crs_h_lego)).set_duration(self.drtn).set_start(0)
        clips.append(text)
        
        text = mpy.TextClip(txt=knowledge,align='West',fontsize=25, font='j:/fonts/yousheTitleHei.ttf',color=clr) \
                .set_fps(25).set_position((self.w*0.1,self.h*crs_h_intro)).set_duration(self.drtn).set_start(0)
        clips.append(text)

        # cover_clip=mpy.CompositeVideoClip(clips)

        # fn=os.path.join(self.src_dir,self.crs_name,self.crs_name+'_building_animation.mp4')
        # cover_clip.write_videofile(fn)

        print('完成')
        return clips

    def build_movie(self,total_secs=54,w=850,h=480):
        print('正在生成主动画……',end='') 
        pics=self.read_pics()
        drtn=total_secs/len(pics)  #总共54秒
        clips=[]
        cover=mpy.ImageClip(os.path.join(self.src_dir,self.crs_name,self.crs_name[4:]+'.jpg')).set_fps(25).set_duration(2).resize((w,h))
        clips.append(cover)
        # clips=ImageSequenceClip(pics,fps=25)
        n=0
        for fn in pics:
            img=mpy.ImageClip(fn).set_fps(25).set_duration(drtn).resize((w,h)).set_start(2+drtn*n)
            clips.append(img)
            n+=1
        
        print('完成')
        return clips

    def put_endvideo(self,clips):
        print('正在加入片尾、logo……',end='')
        endvideo=mpy.VideoFileClip(self.endV,target_resolution=(self.h,self.w)).set_start(56) # 片尾   分辨率是先写h,再写w  大坑
        add_end=clips.append(endvideo)

        logo=mpy.ImageClip(r"I:\\大智小超\\公共素材\\图片类\\00大智小超科学实验室商标.png") \
            .set_fps(25).set_duration(54).resize((80,48)).set_position((740,420)) \
            .set_start(0)
        clips.append(logo)

        finalclip=mpy.CompositeVideoClip(clips)
        print('完成')
        return finalclip


    def put_bgm(self,finalclip,endvideo_duration=4):
        print('正在加入背景音乐……',end='')
        BGM=mpy.AudioFileClip(self.bgm).set_duration(finalclip.duration-endvideo_duration).fx(afx.audio_fadeout,0.8)
        final_audio = mpy.CompositeAudioClip([BGM,finalclip.audio])
        mix=finalclip.set_audio(final_audio)
        print('完成')
        return mix


    def killProcess(self):
        # 处理python程序在运行中出现的异常和错误
        try:
            # pids方法查看系统全部进程
            pids = psutil.pids()
            for pid in pids:
                # Process方法查看单个进程
                p = psutil.Process(pid)
                # print('pid-%s,pname-%s' % (pid, p.name()))
                # 进程名
                if p.name() == 'ffmpeg-win64-v4.1.exe':
                    # 关闭任务 /f是强制执行，/im对应程序名
                    cmd = 'taskkill /f /im ffmpeg-win64-v4.1.exe  2>nul 1>null'
                    # python调用Shell脚本执行cmd命令
                    os.system(cmd)
        except:
            pass
        

    def exp_building_movie(self,exptype='all',total_sec_for_part=25):
        print('正在处理……')
        

        if exptype=='all':
            main_movie=self.build_movie(w=850,h=480)
            cover_text=self.put_cover_text(main_movie)
            add_end_video=self.put_endvideo(cover_text)
            mix=self.put_bgm(add_end_video)
            fn=os.path.join(self.src_dir,self.crs_name,self.crs_name+'_building_animation.mp4')
            0
        elif exptype=='part':
            main_movie=self.build_movie(total_secs=total_sec_for_part,w=1280,h=720)
            cover_text=self.put_cover_text(main_movie)
            mix=mpy.CompositeVideoClip(cover_text)
            fn=os.path.join(self.src_dir,self.crs_name,self.crs_name+'_building_animation_only.mp4')
        else:
            print('无效参数')
            sys.exit(0)

        if self.save_yn=='yes':
            mix.write_videofile(fn)
        
        return mix
        self.killProcess()
        print('All Done')

class AnimationAndVideo:
    # 第二段影片建议在29-44秒之间
    def __init__(self,crs_name='L056陀螺发射器'):
        config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'picToMP4config.txt'))
        self.bgm_src=config['背景音乐']
        self.pic_dir=config['图纸文件夹']
        self.crs_info_src=config['课程信息表']
        self.crs_name=crs_name
        self.end_clip_src=config['片尾']
        self.logo_src=config['logo']

    def read_crs_info(self):
        df_crs_info=pd.read_excel(self.crs_info_src)
        crs_info=df_crs_info[df_crs_info['课程编号']==self.crs_name[0:4]]     
        return crs_info

    def export_mv(self,w=1280,h=720,bgm_src='default'):
        crs_info=self.read_crs_info()

        clip_02_src=os.path.join(self.pic_dir,self.crs_name,self.crs_name+'_clip_02.mp4')
        _clip_02_src=mpy.VideoFileClip(clip_02_src,target_resolution=(h,w)).set_start(0)
        target_sec=56-int(_clip_02_src.duration)

        if _clip_02_src.duration>44:
            print('第二段影片大于44秒，请先剪裁。')
            sys.exit(0)

        building_ani_src=os.path.join('i:\\乐高\\图纸',self.crs_name,self.crs_name+'_building_animation_only.mp4')
        if os.path.exists(building_ani_src):
            print('目录中已存在搭建动画,将用于合并生成视频号影片。')
            clip_01_src=os.path.join(self.pic_dir,self.crs_name,self.crs_name+'_building_animation_only.mp4')
            _clip_01=mpy.VideoFileClip(clip_01_src,target_resolution=(h,w)).set_start(0)
            acc_clip_01 = _clip_01.fl_time(lambda t:  _clip_01.duration/target_sec*t, apply_to=['mask', 'audio'])
            clip_01=acc_clip_01.set_duration(target_sec)
        else:
            print('目录中无搭建动画，正在生成搭建动画序列……')
            building_ani=BuildAnimation(crs_name=self.crs_name,save_yn='no')
            # _clip_01=building_ani.exp_building_movie(exptype='part',total_sec_for_part=target_sec)
            clip_01=building_ani.exp_building_movie(exptype='part',total_sec_for_part=target_sec)

        

        
        # target_sec=10

        # acc_clip_01 = _clip_01.fl_time(lambda t:  _clip_01.duration/target_sec*t, apply_to=['mask', 'audio'])
        # clip_01=acc_clip_01.set_duration(target_sec)

        clip_02=mpy.VideoFileClip(clip_02_src,target_resolution=(h,w)).set_start(clip_01.duration)


        # bg_time=int(clip_01.duration+clip_02.duration)-2*clip_01.duration/_clip_01.duration
        bg_time=int(clip_01.duration+clip_02.duration)-2
        bg=mpy.ColorClip((430,720),color=(0,0,0),ismask=False,duration=bg_time).set_opacity(0.5).set_position((850,0)).set_start(2)

        bg_left=mpy.ColorClip((300,56),color=(51,149,255),ismask=False,duration=bg_time).set_position((275,15)).set_start(2)

        txt_left='科学机器人课'
        txt_title=self.crs_name[4:]
        txt_tool='教具：'+crs_info['教具'].values.tolist()[0]
        txt_big_klg='课程知识点'
        txt_klg=crs_info['知识点'].values.tolist()[0].split('\n')

        clip_left=mpy.TextClip(txt_left,fontsize=40, font='j:/fonts/hongMengHei.ttf',color='#ffffff').set_position((310,18)).set_duration(bg_time).set_start(2)
        clip_title=mpy.TextClip(txt_title,fontsize=54, font='j:/fonts/yousheTitleHei.ttf',color='#ffff00').set_position((int(430/2)-int(len(txt_title)*54/2)+860,22)).set_duration(bg_time).set_start(2)
        clip_tool=mpy.TextClip(txt_tool,fontsize=26, font='j:/fonts/yousheTitleHei.ttf',color='#ffff00').set_position((int(430/2)-int(len(txt_tool)*26/2)+880,110)).set_duration(bg_time).set_start(2)
        clip_big_klg=mpy.TextClip(txt_big_klg,fontsize=46, font='j:/fonts/HYXinHaiXingKaiW.ttf',color='#ffffff').set_position((int(430/2)-int(len(txt_big_klg)*46/2)+850,350)).set_duration(bg_time).set_start(2)
        clip_logo=mpy.ImageClip(self.logo_src).set_fps(25).set_position((20,650)).set_duration(56).resize((110,int(110*253/425))).set_start(0)

        clips=[clip_01,clip_02,bg,clip_logo,clip_title,clip_tool,clip_big_klg]

        for n,text in enumerate(txt_klg):            
            clip_klg=mpy.TextClip(text,fontsize=30, font='j:/fonts/HYXinHaiXingKaiW.ttf',color='#ffffff',align='West').set_position((890,430+n*48)).set_duration(bg_time).set_start(2)
            clips.append(clip_klg)

        clip_end=mpy.VideoFileClip(self.end_clip_src,target_resolution=(h,w)).set_start(bg_time+2)

        clips_rear=[bg_left,clip_left,clip_end]
        clips.extend(clips_rear)
        finalclip=mpy.CompositeVideoClip(clips)

        if bgm_src=='default':
            bgm=mpy.AudioFileClip(self.bgm_src).set_duration(finalclip.duration-clip_end.duration).fx(afx.audio_fadeout,0.8)
        else:
            bgm=mpy.AudioFileClip(bgm_src).set_duration(finalclip.duration-clip_end.duration).fx(afx.audio_fadeout,0.8)
        final_audio = mpy.CompositeAudioClip([bgm,finalclip.audio])
        mix=finalclip.set_audio(final_audio)

        out_mv=os.path.join(self.pic_dir,self.crs_name,self.crs_name+'_视频号.mp4')
        mix.write_videofile(out_mv)

class helpp():
    def __init__(self):
        txt='''
        doLego包括两个类：LegoCons和ConsMovie， 
        LegoCons用于生成html格式的步骤图，并保存为txt格式，
        ConsMovie用于将已有的图片剪辑生成视频。
        
        输入：
        doLego.help.helpConsMove()
        doLego.help.helpLegoCons()
        
        '''
        print(txt)
    
    def helpConsMovie(self):
        txt='''
        doLego.ConsMovie(path,name)
        doLego.ConsMovie.expMovie()
        
        example:
        myMovie=ConsMovie('I:\\乐高\\图纸','002长颈鹿')
        myMovie.expMovie()
        '''
        print(txt)
        
    def helpLegoCons(self):
        txt='''
        doLego.LegoCons(path,name)
        doLego.LegoCons.creatItem()
        doLego.LegoCons.expHTML_lego()
        
        example_1 从已有的图片生成公众号文章后台html格式步骤，并保存为txt文件:
        mylego=LegoCons('I:\\乐高\\图纸','003鳄鱼')
        mylego.expHTML_lego()
        
        example_2 新建一个项目文件夹，该文件夹下将包含“零件总图”、“video”两个子文件夹:
        mylego=LegoCons('I:\\乐高\\图纸','004公鸡')
        mylego.expHTML_lego()
        '''
        print(txt)       
        
    
def run_1(pth,consName,crs_list,s=1):
    if s==1: #仅新建项目
        mylego=LegoCons(pth,consName)
    elif s==2:#改名及输出列表
        mylego=LegoCons(pth,consName)        
        mylego.creatItem()
        mylego.makeList()
        mylego.putTag()
    elif s==3: #制作影片
        mylego=LegoCons(pth,consName)        
        # mylego.creatItem()
        # mylego.makeList()
        # mylego.putTag()
        mylego.putTag_wedo_coverpage()
        mylego.expHTML_lego()    
        myMovie=ConsMovie(pth,consName)
        myMovie.exportMovie()
    
def run_export_Poster():
    myPoster=LegoWeekly('I:\大智小超\宣传\搭建主题')
    myPoster.expPoster()
    
def run_test():
    mylego=LegoCons('I:\\乐高\\图纸','019露营的折叠椅')
    mylego.putTag_wedo_coverpage()

if __name__=='__main__':
    #   run_1('I:\\乐高\\图纸','006鸭子',2)
    #     run_1('I:\\乐高\\图纸','021天平',2)
    # run_1('I:\\乐高\\图纸','035啃骨头的小狗',3)
    #     run_export_Poster()
    #     run_test()

    # my=BuildAnimation('L033双翼飞机')
    # my.exp_building_movie(exptype='part')

    my=AnimationAndVideo(crs_name='L056陀螺发射器')
    my.export_mv(w=1280,h=720,bgm_src='default')
    # k=my.read_crs_info()
    # p=k['知识点'].values.tolist()[0]
    # print(p.split('\n'))
