import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'LegoStudentPic'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoCrstoPPT'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoPosterAfterClass'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'picToMp4'))
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

def merge_animation_mv(crs_name='L056陀螺发射器',method_merge=1):
    # 第二段影片建议在29-44秒之间            
    try:
        if method_merge==1:
            building_ani_src=os.path.join('i:\\乐高\\图纸',crs_name,crs_name+'_building_animation_only.mp4')
            if not os.path.exists(building_ani_src):
                build_ani=picToMp4.BuildAnimation(crs_name=crs_name)
                build_ani.exp_building_movie(exptype='part')  #part仅导出动画，无音乐， all导出完成影片。
            my=picToMp4.AnimationAndVideo(crs_name=crs_name)
            my.export_mv(w=1280,h=720)
        elif method_merge==2:
            my=picToMp4.AnimationAndVideo(crs_name=crs_name)
            my.export_mv(w=1280,h=720)
        else:
            print('无此参数')
    except:
        print('出错。请检查目录中是否有相应文件。')

def stdpicWhiteMark():
    pic=LegoStdPicMark.pics()
    pic.putCover(height=2250)

def makePpt(crsName='L037认识零件',copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程'):
    mypics=legoCrstoPPT.picToPPT(crsName)
    mypics.ExpPPT(copyToCrsDir=copyToCrsDir,crsPPTDir=crsPPTDir)

def legoPoster(crsName='L035啃骨头的小狗',crsDate=20201020):
    weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通日期计算星期
    my=posterAfterClass(weekday=weekday)
    #     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw(crsName,crsDate)    

def picsDistribute(crsDate,crsName):
    weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通日期计算星期
    stu_pics=LegoStudentPicDistribute.LegoPics(crsDate,crsName,weekday)
    stu_pics.dispatch()

def pics_distribute_and_make_poster(place='超智幼儿园',crsDate=20200929,crsName='L033双翼飞机',TeacherSig='阿晓老师'):
    weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通日期计算星期
    crsName_distibute=str(crsDate)+crsName[4:]
    picsDistribute(crsDate=crsDate,crsName=crsName)
    my=posterAfterClass(weekday=weekday,place=place)  
    my.PosterDraw(crs_nameInput=crsName,dateInput=crsDate,TeacherSig=TeacherSig)


#将步骤图生成1分钟视频放上视频号
# makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='L056陀螺发射器',crs_list='课程信息表.xlsx',dealtype='makeMovie',src='ldcad')

#将步骤图生成搭建视频，与原拍摄视频合并  拍摄mp4建议29-44秒
merge_animation_mv(crs_name='L053慢悠悠的大象',method_merge=1) #方法1：先生成搭建动画，保存后再生成影片，不容易有黑色卡顿， 方法2：直接通过Png生成动画。

#将步骤图导出PPT
# makePpt('L046圣诞老人来了',copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程')

#学期末为照片加上灰背景及知识点等
# stdpicWhiteMark()

# #按名字分配照片，并生成课后发给家长的照片：
# pics_distribute_and_make_poster(place='超智幼儿园',crsDate=20201215,crsName='L056陀螺发射器',TeacherSig='阿晓老师')
