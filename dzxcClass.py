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

def makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='L017毛毛虫',crs_list='课程信息表.xlsx',dealtype='exportList'):
    consName=crsName
    if dealtype=='exportList':
        mylego=picToMp4.LegoCons(pth,consName)        
        mylego.creatItem()
        mylego.makeList()
        mylego.putTag()
    elif dealtype=='makeMovie':
        mylego=picToMp4.LegoCons(pth,consName)      
        mylego.creatItem()
        mylego.makeList()
        mylego.putTag()
        mylego.putTag_wedo_coverpage()
        mylego.expHTML_lego()    
        myMovie=picToMp4.ConsMovie(pth,consName,crs_list)
        myMovie.exportMovie()

def stdpicWhiteMark():
    pic=LegoStdPicMark.pics()
    pic.putCover(height=2250)

def exportPpt(crsName='L037认识零件',copyToCrsDir='yes',crsPPTDir='I:\\乐高\\乐高WeDo\\课程'):
    mypics=legoCrstoPPT.picToPPT(crsName)
    mypics.ExpPPT(copyToCrsDir=copyToCrsDir,crsPPTDir=crsPPTDir)

def legoPoster(crsName='L035啃骨头的小狗',crsDate=20201020,weekday=2):
    my=posterAfterClass(weekday=weekday)
    #     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw(crsName,crsDate)    

def picsDistribute(crsDate,crsName,weekday):
    stu_pics=LegoStudentPicDistribute.LegoPics(crsDate,crsName,weekday)
    stu_pics.dispatch()

def pic_distribute_and_exp_poster(place='超智幼儿园',crsDate=20200929,crsName='L033双翼飞机',weekday=6,TeacherSig='阿晓老师'):
    crsName_distibute=str(crsDate)+crsName[4:]
    picsDistribute(crsDate=crsDate,crsName=crsName,weekday=weekday)
    my=posterAfterClass(weekday=weekday,place=place)  
    my.PosterDraw(crs_nameInput=crsName,dateInput=crsDate,TeacherSig=TeacherSig)


#将步骤图生成1分钟视频放上视频号
# makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='L017毛毛虫',crs_list='课程信息表.xlsx',dealtype='makeMovie')

#将步骤图导出PPT
# exportPpt('L038旋转飞椅',copyToCrsDir='yes',crsPPTDir='I:\\乐高\\乐高WeDo\\课程')

#学期末为照片加上灰背景及知识点等
# stdpicWhiteMark()

#按名字分配照片，并生成课后发给家长的照片：
pic_distribute_and_exp_poster(place='超智幼儿园',crsDate=20201024,crsName='L037认识零件',weekday=6,TeacherSig='阿晓老师')

