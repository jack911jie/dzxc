import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'LegoStudentPic'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoCrstoPPT'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'legoPosterAfterClass'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'picToMp4'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'poster'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'TermSummary'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'query'))
sys.path.append(os.path.join(os.path.dirname(os.path.realpath(__file__)),'module'))
import Summary_hd
import BeforeClass
import picToMp4
import LegoStudentPicDistribute
from legoPoster import poster as posterAfterClass
import LegoStdPicMark
import legoCrstoPPT
import QueryForScore
import QueryForClassTaken
import iptcinfo3
from datetime import datetime

def view_score(place_input='001-超智幼儿园',std_name='w401',sort_by='剩余积分',plus_tiyan='no'):
    myQuery=QueryForScore.query(place_input=place_input)
    res=myQuery.query_for_scores(std_name,plus_tiyan=plus_tiyan)
    res.sort_values(by=sort_by,inplace=True)
    print(res)
    # print(datetime.now().strftime('%Y-%m-%d'))

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

def stdpicWhiteMark(height=2250,term='2020秋',crop='yes',bigger='yes',weekday=[2]):
    pic=LegoStdPicMark.pics()
    for wd in weekday:
        pic.putCover(height=height,term=term,crop=crop,bigger=bigger,weekday=wd)

def stdpicYellowMark(place_input='001-超智幼儿园',term='2021春',weekdays=[1,4],start_date='20210103',end_date='20210506'):
    pic=LegoStdPicMark.SimpleMark(place_input=place_input)
    pic.put(term=term,weekdays=weekdays,start_date=start_date,end_date=end_date)

def makePpt(crsName='L037认识零件',copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程',pos_pic='no',lxfml_mode='new'):    
    mypics=legoCrstoPPT.picToPPT(crsName)
    if pos_pic=='yes':        
        mypics.ExpPPT(copyToCrsDir=copyToCrsDir,lxfml_mode=lxfml_mode,crsPPTDir=crsPPTDir)
        mypics.inner_box_pos(save='yes',lxfml_mode=lxfml_mode)
    elif pos_pic=='pos_pic_only':
        mypics.inner_box_pos(save='yes',lxfml_mode=lxfml_mode)
    elif pos_pic=='no':
        mypics.ExpPPT(copyToCrsDir=copyToCrsDir,lxfml_mode=lxfml_mode,crsPPTDir=crsPPTDir)
    else:
        mypics.ExpPPT(copyToCrsDir=copyToCrsDir,lxfml_mode=lxfml_mode,crsPPTDir=crsPPTDir)

def legoPoster(crsName='L035啃骨头的小狗',crsDate=20201020,mode=''):
    weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通日期计算星期
    my=posterAfterClass(weekday=weekday,mode=mode)
    #     my.PosterDraw('可以伸缩的夹子')      
    my.PosterDraw(crsName,crsDate)    

def picsDistribute(crsDate,place,crsName,term,force_weekday=0,mode=''):
    if force_weekday==0:
        weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通过日期计算星期
    else:
        weekday=force_weekday
    stu_pics=LegoStudentPicDistribute.LegoPics(crsDate,crsName,place,weekday,term,mode=mode)
    stu_pics.dispatch()

def pics_distribute_and_make_poster(place='001-超智幼儿园',term='2021春',crsDate_name='20210419-L066弹力小车',force_weekday=0,TeacherSig='阿晓老师',pic_forced_ht='',copy_to_feedback_dir='no',mode=''):
    crsDate=crsDate_name.split('-')[0]
    crsName=crsDate_name.split('-')[1]
    if force_weekday==0:
        weekday=datetime.strptime(str(crsDate),'%Y%m%d').weekday()+1 #通过日期计算星期
    else:
        weekday=force_weekday
    crsName_distibute=str(crsDate)+crsName[4:]
    picsDistribute(crsDate=crsDate,place=place,crsName=crsName,term=term,force_weekday=force_weekday,mode=mode)
    # print(crsDate_name[10:])
    pic_to_feedback=LegoStudentPicDistribute.LegoPics(crsDate=crsDate_name[0:8],crsName=crsDate_name[9:],weekday=weekday,term=term,mode='')
    pic_to_feedback.dispatch_after_class()
    pics=LegoStdPicMark.AfterClassMark()
    pics.group_put(crs_date=crsDate_name[0:8],crs_name=crsDate_name[9:],forced_ht=pic_forced_ht)
    my=posterAfterClass(weekday=weekday,term=term,place_input=place,mode=mode)  
    my.PosterDraw(crs_nameInput=crsName,dateInput=crsDate,TeacherSig=TeacherSig,copy_to_feedback_dir=copy_to_feedback_dir)

def before_class_poster(crsDate_name='20210619-L094游泳的鲨鱼',time_crs_input='1100-1230'):
    date_crs_input=crsDate_name.split('-')[0]
    crs_name_input=crsDate_name.split('-')[1]
    my=BeforeClass.LegoClass()
    my.exp_poster(date_crs_input=date_crs_input,time_crs_input=time_crs_input,crs_name_input=crs_name_input)

def std_grow_book(std_name='韦宇浠',start_date='20200922',end_date='20210309',weekday='2',term='2020秋',tch_name='阿晓老师',k=1.25):
    my=Summary_hd.data_summary()
    # my.rose(std_name=std_name,weekday=weekday)
    my.exp_poster(std_name=std_name,start_date=start_date,end_date=end_date,weekday=weekday,term=term,tch_name=tch_name,k=k)

def std_ability_rose(std_name='韦成宇',term='2020秋',weekday='6'):
    my=Summary_hd.data_summary()
    pic=my.rose(std_name=std_name,xls='E:\\WXWork\\1688852895928129\\WeDrive\\大智小超科学实验室\\001-超智幼儿园\\每周课程反馈\\2021春-学生课堂学习情况反馈表（周六）.xlsx')
    # my.exp_poster(std_name=std_name,start_date=start_date,end_date=end_date,weekday=weekday,term=term,tch_name=tch_name,k=k)

def stage_report(std_names=['韦华晋','黄建乐'],start_date='20210301',end_date='20210801', \
                    cmt_date='20210719',tb_list=[['2021春','w4']], \
                    tch_name=['阿晓','杨芳芳'],mode='all',k=1):
    stds=Summary_hd.data_summary()
    for n in std_names:
        stds.exp_a4_16(std_name=n,start_date=start_date,end_date=end_date, \
                    cmt_date=cmt_date,tb_list=tb_list, \
                    tch_name=tch_name,mode=mode,k=k)

def tiyanke_step_mark(height=2250,term='2021秋',crop='yes',bigger='yes',weekday=5):
    p=LegoStdPicMark.StepMark()
    p.putCover(height=height,term=term,crop=crop,bigger=bigger,weekday=weekday)

def check_crs_dup(place_input='001-超智幼儿园',term='2021秋',weekday=5,fn='c:/Users/jack/desktop/w5待排课程.txt',show_res='yes',write_file='no'):
    qry=QueryForClassTaken.Query(place_input=place_input)
    # qry.std_class_taken(weekday=[1,4],display='print_list',format='only_clsnam')
    # qry.check_duplicate(term='2021秋',weekday=1,crs_name='L026跷跷板') 
    conflict=qry.check_conflict(term=term,weekday=weekday,fn=fn,show_res=show_res,write_file=write_file)


if __name__=='__main__':
#排课检查课程冲突
    # check_crs_dup(place_input='001-超智幼儿园',term='2021秋',weekday=6,fn='c:/Users/jack/desktop/待排课程.txt',show_res='yes',write_file='no')

#体验课给照片打标
    # tiyanke_step_mark(height=2250,term='2021秋',crop='yes',bigger='yes',weekday=1)

#将步骤图生成1分钟视频放上视频号
    # makeLegoConsMovie(pth='I:\\乐高\\图纸',crsName='L056陀螺发射器',crs_list='课程信息表.xlsx',dealtype='makeMovie',src='ldcad')

#将步骤图生成搭建视频，与原拍摄视频合并  拍摄mp4建议29-44秒
    #方法1：先生成搭建动画，保存后再生成影片，不容易有黑色卡顿， 方法2：直接通过Png生成动画。 背景音乐默认参数为default，可输入mp3文件地址替换默认背景音乐。
    # merge_animation_mv(crs_name='L046圣诞老人来了',method_merge=1,bgm_src='e:/temp/JingleBells2.mp3') 


#学期末为照片加上灰背景及知识点等
    # stdpicWhiteMark(height=2250,term='2021春',crop='yes',bigger='yes',weekday=[1,4])

#学期末为照片加上黄色背景
    # stdpicYellowMark(place_input='001-超智幼儿园',term='2021春',weekdays=[4],start_date='20210303',end_date='20210806')

#16节课/或阶段课程学习报告
    # list_w4=['w4',['陈锦媛','陆炳辰','刘嘉祥','李贤斌','庞孙钜','沈金辰','陶盛挺','韦宇浠','李家捷','林祖葳']]
    # list_w1=['w1',['邓立文','黄昱涵','农雨蒙','覃熙雅','廖韦朗','王丹亭','李宛晏','韦万祎','谢威年','黄进桓']]
    # list_all=[list_w1,list_w4]
    # for list_cus in list_all:
    #     stage_report(std_names=list_cus[1],start_date='20210301',end_date='20210801', \
    #                         cmt_date='20210719',tb_list=[['2021春',list_cus[0]]], \
    #                         tch_name=['阿晓','芳芳'],mode='all',k=1)

    # stage_report(std_names=['廖韦朗'],start_date='20210301',end_date='20210801', \
    #                         cmt_date='20210719',tb_list=[['2021春','w1']], \
    #                         tch_name=['阿晓','芳芳'],mode='all',k=1)

#学员成长手册
    # std_grow_book(std_name='韦华晋',start_date='20200922',end_date='20210309',weekday='2',term='2020秋',tch_name='阿晓老师')

#学员能力玫瑰图
    # std_ability_rose(std_name='韦成宇',term='2020秋',weekday='6')

#将步骤图导出PPT
    # makePpt('L112飞行的小鸟',copyToCrsDir='yes',crsPPTDir='I:\\乐高\\乐高WeDo\\课程',pos_pic='yes',lxfml_mode='new')

# #查看积分
    # view_score(place_input='001-超智幼儿园',std_name='w601',sort_by='剩余积分',plus_tiyan='no')

# #按名字分配照片，并生成课后发给家长的照片：
    pics_distribute_and_make_poster(place='001-超智幼儿园', term='2021秋',crsDate_name='20211018-L111可爱的小蜜蜂', 
        force_weekday=0,TeacherSig='阿晓老师',pic_forced_ht=1200,copy_to_feedback_dir='yes',mode='')  

# 课前生成海报
    # before_class_poster(crsDate_name='20211023-L112飞行的小鸟',time_crs_input='1100-1230')

 