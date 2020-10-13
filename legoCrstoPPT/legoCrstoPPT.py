import os
import platform
import numpy as np
import re
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm,Pt,Cm
from pptx.dml.color import RGBColor

class picToPPT:
    def __init__(self,picSrc):
        if platform.system().lower()=='linux':
            self.template='/home/jack/data/乐高/图纸/template.pptx'
            self.wtmark='/home/jack/data/乐高/图纸/大智小超水印.png'
        else:
            self.template='i:/乐高/图纸/template.pptx'
            self.wtmark='i:/乐高/图纸/大智小超水印.png'
        self.picSrc=picSrc
    
    def ExpPPT(self):
        print('\n正在处理……',end='')       
    
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
                if ptn_block_list.match(blk) or ptn_block_list2.match(blk):
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
            step_blkList=pd.read_excel(os.path.join(self.picSrc,self.picSrc.split('/')[-1]+'-步骤零件.xlsx')).replace(np.nan,'')['零件名称'].tolist()
#             print(step_blkList)
            
            
            prs=Presentation(self.template)            
            left=Cm(0)
            top=Cm(1.4)
            height=Cm(17.69)
            
#             left_wtmk=Cm(5)
#             top_wtmk=Cm(5)
            left_wtmk=Cm(0)
            top_wtmk=Cm(1.4)
            
            for i,img in enumerate(picList[0]):
                blank_slide_layout=prs.slide_layouts[1]
                slide=prs.slides.add_slide(blank_slide_layout)
                pic=slide.shapes.add_picture(img,left,top,height=height)
                textbox=slide.shapes.add_textbox(Cm(2),Cm(5),Cm(5),Cm(2.5))
                p = textbox.text_frame.add_paragraph()
                
                if i>=picList[1]:
#                     print(i,i-picList[1])
                    try:
                        p.text=step_blkList[i-picList[1]]
                        p.font.size=Pt(30)
                        p.font.color.rgb = RGBColor(22, 56, 153)
                    except:
                        pass
                
                pic_wtmark=slide.shapes.add_picture(self.wtmark,left_wtmk,top_wtmk) #加水印
                
            newFn=os.path.join(self.picSrc,self.picSrc.split('/')[-1]+'.pptx')
            prs.save(newFn)
            print('完成，文件名：{}'.format(newFn))                              
            
        picList=picList()
        picToPPT(picList)          
        
if __name__=='__main__':
    mypics=picToPPT('i:/乐高/图纸/034夏天的手摇风扇')
#     mypics=picToPPT('/home/jack/data/乐高/图纸/031回力赛车')
    mypics.ExpPPT()