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

def makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='L017毛毛虫',crs_list='课程信息表.xlsx',dealtype='makeList'):
    consName=crsName
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
# makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='N001电钻',crs_list='课程信息表.xlsx',dealtype='makeMovie')


#将步骤图导出PPT
# makePpt('L046圣诞老人来了',copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程')

#学期末为照片加上灰背景及知识点等
# stdpicWhiteMark()

#按名字分配照片，并生成课后发给家长的照片：
pics_distribute_and_make_poster(place='超智幼儿园',crsDate=20201114,crsName='L035啃骨头的小狗',TeacherSig='阿晓老师')