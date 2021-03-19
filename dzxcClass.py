import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'LegoStudentPic'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoCrstoPPT'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoPosterAfterClass'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'picToMp4'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'poster'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'TermSummary'))
import Summary_hd
import BeforeClass
import picToMp4
import LegoStudentPicDistribute
from legoPoster import poster as posterAfterClass
import LegoStdPicMark
import legoCrstoPPT
import iptcinfo3
from datetime import datetime

def makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='L017毛毛虫',crs_list='课程信息表.xlsx',dealtype='makeList',src='ldcad'):
    consName=crsName
    if src=='ldcad':
        my=picToMp4.BuildAnimation(crs_name=consName)
        my.exp_building_movie(exptype='part')  #part仅导出动画，无音乐， all导出完成影片。
    else:
        if dealtype=='makeList':
            mylego=picToMp4.LegoCons(pth,consName)        
            mylego.creatItem()
            mylego.makeList()
            mylego.putTag()
        elif dealtype=='makeMovie':
            mylego=picToMp4.LegoCons(pth,consName)      
            mylego.creatItem()
            mylego.makeList()
            mylego.putTag()
            # mylego.putTag_wedo_coverpage() #已不需要这步，第一张片头的字幕在moviepy中添加，不需另外生成一张图片。
            mylego.expHTML_lego()    
            myMovie=picToMp4.ConsMovie(pth,consName)
            myMovie.exportMovie()

def merge_animation_mv(crs_name='L056陀螺发射器',method_merge=1,bgm_src='default'):
    # 第二段影片建议在29-44秒之间            
    try:
        if method_merge==1:
            building_ani_src=os.path.join('i:\\乐高\\图纸',crs_name,crs_name+'_building_animation_only.mp4')
            if not os.path.exists(building_ani_src):
                build_ani=picToMp4.BuildAnimation(crs_name=crs_name)
                build_ani.exp_building_movie(exptype='part')  #part仅导出动画，无音乐， all导出完成影片。
            my=picToMp4.AnimationAndVideo(crs_name=crs_name)
            my.export_mv(w=1280,h=720,bgm_src=bgm_src)
        elif method_merge==2:
            my=picToMp4.AnimationAndVideo(crs_name=crs_name)
            my.export_mv(w=1280,h=720,bgm_src=bgm_src)
        else:
            print('无此参数')
    except:
        print('出错。请检查目录中是否有相应文件。')

def stdpicWhiteMark(height=2250,weekday=[2]):
    pic=LegoStdPicMark.pics()
    for wd in weekday:
        pic.putCover(height=height,weekday=wd)

def makePpt(crsName='L037认识零件',copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程'):
    mypics=legoCrstoPPT.picToPPT(crsName)
    mypics.ExpPPT(copyToCrsDir=copyToCrsDir,crsPPTDir=crsPPTDir)

def legoPoster(crsName='L035啃骨头的小狗',crsDate=20201020):
    weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通日期计算星期
    my=posterAfterClass(weekday=weekday)
    #     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw(crsName,crsDate)    

def picsDistribute(crsDate,place,crsName,term):
    weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通日期计算星期
    stu_pics=LegoStudentPicDistribute.LegoPics(crsDate,crsName,place,weekday,term)
    stu_pics.dispatch()

def pics_distribute_and_make_poster(place='5-超智幼儿园',term='2021春',crsDate=20200929,crsName='L033双翼飞机',TeacherSig='阿晓老师'):
    weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通日期计算星期
    crsName_distibute=str(crsDate)+crsName[4:]
    picsDistribute(crsDate=crsDate,place=place,crsName=crsName,term=term)
    my=posterAfterClass(weekday=weekday,term=term,place_input=place)  
    my.PosterDraw(crs_nameInput=crsName,dateInput=crsDate,TeacherSig=TeacherSig)

def before_class_poster(date_crs_input='20210306',time_crs_input='1100-1230',crs_name_input='L063汽车雨刮器'):
    my=BeforeClass.LegoClass()
    my.exp_poster(date_crs_input=date_crs_input,time_crs_input=time_crs_input,crs_name_input=crs_name_input)

def std_grow_book(std_name='韦宇浠',start_date='20200922',end_date='20210309',weekday='2',term='2020秋',tch_name='阿晓老师',k=1.25):
    my=Summary_hd.data_summary()
    # my.rose(std_name=std_name,weekday=weekday)
    my.exp_poster(std_name=std_name,start_date=start_date,end_date=end_date,weekday=weekday,term=term,tch_name=tch_name,k=k)

def std_ability_rose(std_name='韦宇浠',weekday='2'):
    my=Summary_hd.data_summary()
    my.rose(std_name=std_name,weekday=weekday)
    # my.exp_poster(std_name=std_name,start_date=start_date,end_date=end_date,weekday=weekday,term=term,tch_name=tch_name,k=k)


#将步骤图生成1分钟视频放上视频号
# makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='L056陀螺发射器',crs_list='课程信息表.xlsx',dealtype='makeMovie',src='ldcad')

#将步骤图生成搭建视频，与原拍摄视频合并  拍摄mp4建议29-44秒
#方法1：先生成搭建动画，保存后再生成影片，不容易有黑色卡顿， 方法2：直接通过Png生成动画。 背景音乐默认参数为default，可输入mp3文件地址替换默认背景音乐。
# merge_animation_mv(crs_name='L046圣诞老人来了',method_merge=1,bgm_src='e:/temp/JingleBells2.mp3') 
#
#将步骤图导出PPT
# makePpt('L068厉害的投掷器',copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程')

#学期末为照片加上灰背景及知识点等
# stdpicWhiteMark(height=2250,weekday=[2,6])

#学员成长手册
# std_grow_book(std_name='韦宇浠',start_date='20200922',end_date='20210309',weekday='2',term='2020秋',tch_name='阿晓老师')

#学员能力玫瑰图
# std_ability_rose(std_name='韦宇浠',weekday='2')

# #按名字分配照片，并生成课后发给家长的照片：
pics_distribute_and_make_poster(place='5-超智幼儿园',term='2021春',crsDate=20210320,crsName='L068厉害的投掷器',TeacherSig='阿晓老师')

# 课前生成海报
# before_class_poster(date_crs_input='20210320',time_crs_input='1100-1230',crs_name_input='L068厉害的投掷器')
