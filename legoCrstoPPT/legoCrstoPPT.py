import os
import sys
sys.path.append('i:/py/dzxc/module')
from readConfig import readConfig
import platform
import numpy as np
import re
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm,Pt,Cm
from pptx.dml.color import RGBColor
from openpyxl import Workbook,load_workbook
from shutil import copyfile,copytree

class picToPPT:
    def __init__(self,crsName):
        self.crsName=crsName
        config=readConfig(os.path.join(os.path.dirname(os.path.realpath(__file__)),'crstoPPT.config'))

        if platform.system().lower()=='linux':
            self.picDir='/home/jack/data/乐高/图纸'
            self.template='/home/jack/data/乐高/图纸/template.pptx'
            self.wtmark='/home/jack/data/乐高/图纸/大智小超水印.png'
        else:
            self.picDir=config['图纸文件夹']
            self.template=config['PPT模板']
            self.wtmark=config['PPT水印']
            self.legoCodeList=config['乐高零件编号']
        self.picSrc=os.path.join(self.picDir,crsName)

    def makeDirs(self,dirName):
        if not os.path.exists(dirName):
            print('新项目，正在创建文件夹……',end='')
            os.makedirs(dirName)
            os.makedirs(os.path.join(dirName,'零件总图'))
            os.makedirs(os.path.join(dirName,'video'))
            print('创建完成,请在文件夹中放入步骤图片。')
            sys.exit(0)
        else:
            print('已有项目……',end='')
            files=os.listdir(dirName)
            if len(files)<3:
                print('请先做好步骤图片')
                sys.exit (0)

    def test_stepXls(self):
        xlsName=os.path.join(self.picSrc,self.crsName+'-ppt步骤零件名称.xlsx')
        if not os.path.exists(xlsName):
            blNames=self.blockNames()
            wb=Workbook()
            tb=wb.active
            tb['A1']='序号'
            tb['B1']='零件名称'
            tb['D1']='零件名称（识别）'
            row=2
            for file in os.listdir(self.picSrc):
                ptn='\d{3}_\dx.png'                
                if re.match(ptn,file):
                    tb['A'+str(row)]=row-1
                    row+=1
            
            for k,blk in enumerate(blNames):
                tb['D'+str(k+2)]=blk

            wb.save(xlsName)
            print('新建了步骤文件，请在文件中先输入零件名称。')
            sys.exit(0)
        else:
            df=pd.read_excel(xlsName)
            blockNum=df['零件名称'].count()
            if blockNum==0:
                print('未输入零件名称，请先写入。')
                sys.exit(0)

    def renameFiles(self):
        print('正在查看是否需要修改文件名……',end='')
        for file in os.listdir(self.picSrc):
            if file[-3:].lower()=='png':
                if len(file)<10:
                    oldname=os.path.join(self.picSrc,file)
                    newname=os.path.join(self.picSrc,file.zfill(10))
                    os.rename(oldname,newname)
        print('完成')

    def copytoCrsDir(self,crsPPTDir='I:\\乐高\\乐高WeDo\\课程'):
        desDir=os.path.join(crsPPTDir,self.crsName)
        if not os.path.exists(desDir):
            os.makedirs(desDir)
        
        oldName_video=os.path.join(self.picSrc,'video')
        oldName_ppt=os.path.join(self.picSrc,self.crsName+'_00.pptx')
        newName_video=os.path.join(desDir,'video')
        newName_ppt=os.path.join(desDir,self.crsName+'.pptx')

        copyfile(oldName_ppt,newName_ppt)
        if not os.path.exists(newName_video):
            copytree(oldName_video,newName_video)
        os.startfile(desDir)
        os.startfile(newName_ppt)

    def blockNames(self):
        fn=os.path.join(self.picDir,self.crsName,self.crsName[4:]+'.lxfml')
        bl=pd.read_excel(self.legoCodeList)
        bl['编号'].astype('object')
        with open (fn,'r',encoding='utf-8') as f:
            rl=f.readlines()

        ptn1='(?<=\<Part refID\=\"\d{1}\" designID\=\")(.*)(?=\;A\" materials\=)'
        ptn2='(?<=\<Part refID\=\"\d{2}\" designID\=\")(.*)(?=\;A\" materials\=)'

        a=[]
        for line in rl:
            p=re.findall(ptn1,line)
            if p:
                a.append(p[0])

            p=re.findall(ptn2,line)
            if p:
                a.append(p[0])

        out=[]
        for  aa in a:
            try:
                aa=int(aa)
            except:
                pass
            try:
                blkName=bl[bl['编号']==aa]['中文名称'].values.tolist()[0]
                out.append(blkName)
            except:
                out.append(aa)
        return out

    def ExpPPT(self,copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程'):
        print('\n正在处理：')   
        self.makeDirs(self.picSrc)
        self.renameFiles()    
        self.test_stepXls()    
        def picList():
            ptn_block_list=re.compile(r'\d*_1x_tagged.png')
            ptn_block_list2=re.compile(r'\d*_1x.png')
            for blk_old in os.listdir(self.picSrc):
                if blk_old[-3:].lower()=='png' and len(blk_old)<10:
                    os.rename(os.path.join(self.picSrc,blk_old),os.path.join(self.picSrc,blk_old.zfill(10)))
                
            for blk_old2 in os.listdir(self.picSrc):
                if blk_old2[-3:].lower()=='png' and len(blk_old2)<10:
                    os.rename(os.path.join(self.picSrc,'零件总图',blk_old2),os.path.join(self.picSrc,'零件总图',blk_old2.zfill(10)))
            
            pic_steps=[]
            for blk in os.listdir(os.path.join(self.picSrc,'零件总图')):
                # if ptn_block_list.match(blk) or ptn_block_list2.match(blk):
                if ptn_block_list2.match(blk): #找出零件总图的数量
                    pic_steps.append(os.path.join(self.picSrc,'零件总图',blk))
            pic_steps.sort()
            summary_num=len(pic_steps)
        
            ptn=r'\d.*png'
            picList=[]
            for fn in os.listdir(self.picSrc):
                if re.match(ptn,fn):
                    picList.append(os.path.join(self.picSrc,fn))
            picList.sort()
            #             print(picList)

            pic_steps.extend(picList)
    
            #             print(pic_steps)
    
            return [pic_steps,summary_num]

        def picToPPT(picList):
            step_blkList=pd.read_excel(os.path.join(self.picSrc,self.crsName+'-ppt步骤零件名称.xlsx')).replace(np.nan,'')['零件名称'].tolist()
            #             print(step_blkList)
            
            
            prs=Presentation(self.template)            
            left=Cm(0)
            top=Cm(1.4)
            height=Cm(17.69)
            
            #             left_wtmk=Cm(5)
            #             top_wtmk=Cm(5)
            left_wtmk=Cm(0)
            top_wtmk=Cm(1.4)

            lowerPosList_1=['2x16单位黑色板','12单位绿色梁','12单位绿色孔砖','1×16梁']
            lowerPosList_2=['主机','电机']

            blank_slide_layout=prs.slide_layouts[1]
            slide=prs.slides.add_slide(blank_slide_layout)
            picTitle=slide.shapes.add_picture(os.path.join(self.picSrc,self.crsName[4:]+'.jpg'),left,top,height=height) #加入封面图
            for i,img in enumerate(picList[0]):                
                slide=prs.slides.add_slide(blank_slide_layout)
                pic=slide.shapes.add_picture(img,left,top,height=height)
                
                
                if i>=picList[1]: #跳过零件总图的数量
                #                     print(i,i-picList[1])
                    try:                        
                        txt=step_blkList[i-picList[1]] #根据文件内容决定广西框的位置下移程度，以免遮挡零件图片
                        if txt in lowerPosList_1:
                            y_textbox=Cm(6)
                        elif txt in lowerPosList_2:
                            y_textbox=Cm(7)
                        else:
                            y_textbox=Cm(5)
                        textbox=slide.shapes.add_textbox(Cm(2),y_textbox,Cm(5),Cm(2.5))
                        p = textbox.text_frame.add_paragraph()
                        p.line_spacing=1.2 #行间距


                        p.text=txt
                        p.font.size=Pt(24)
                        p.font.color.rgb = RGBColor(22, 56, 153)
                    except:
                        pass
                
                pic_wtmark=slide.shapes.add_picture(self.wtmark,left_wtmk,top_wtmk) #加水印
                
            newFn=os.path.join(self.picSrc,self.crsName+'_00.pptx')
            prs.save(newFn)

            if copyToCrsDir=='yes':
                self.copytoCrsDir(crsPPTDir=crsPPTDir)

            print('ppt已导出完成，文件名：{}'.format(newFn))                              
            
        picList=picList()
        picToPPT(picList)          
        
if __name__=='__main__':
    mypics=picToPPT('L038旋转飞椅')
    # print(mypics.blockNames())
#     mypics=picToPPT('/home/jack/data/乐高/图纸/031回力赛车')
    mypics.ExpPPT(copyToCrsDir='no',crsPPTDir='I:\\乐高\\乐高WeDo\\课程')
    # mypics.makeDirs()
    # mypics.copytoCrsDir()