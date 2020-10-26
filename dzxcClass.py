import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'LegoStudentPic'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoCrstoPPT'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoPosterAfterClass'))
import doLego
import LegoStudentPicDistribute
from legoPoster import poster as posterAfterClass
import LegoStdPicMark
import legoCrstoPPT

def dolego(crsName='037认识零件',dealtype=1):
    doLego.run_1('I:\\乐高\\图纸',crsName,'课程信息表.xlsx',dealtype)

def stdpicWhiteMark():
    pic=LegoStdPicMark.pics()
    pic.putCover(height=2250)

def exportPpt(crsName='037认识零件',copyToCrsDir='yes',crsPPTDir='I:\\乐高\\乐高WeDo\\课程'):
    mypics=legoCrstoPPT.picToPPT(crsName)
    mypics.ExpPPT(copyToCrsDir=copyToCrsDir,crsPPTDir=crsPPTDir)

def legoPoster(crsName='035啃骨头的小狗',crsDate=20201020,weekday=2):
    my=posterAfterClass(weekday=weekday)
    #     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw(crsName,crsDate)    

def picsDistribute(crsName,weekday):
    stu_pics=LegoStudentPicDistribute.LegoPics(crsName,weekday)
    stu_pics.dispatch()

def pic_distribute_and_exp_poster(crsDate=20201024,crsName='037认识零件',weekday=6):
    crsName_distibute=str(crsDate)+crsName[3:]
    picsDistribute(crsName=crsName_distibute,weekday=weekday)
    legoPoster(crsName=crsName,crsDate=crsDate,weekday=weekday)

# dolego('038旋转飞椅',2)
exportPpt('038旋转飞椅',copyToCrsDir='yes',crsPPTDir='I:\\乐高\\乐高WeDo\\课程')
# stdpicWhiteMark()
# pic_distribute_and_exp_poster(crsDate=20201024,crsName='037认识零件',weekday=6)

