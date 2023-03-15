import os
import sys
sys.path.append('i:/py/modules')
import composing
import json
import pandas as pd
from PIL import Image,ImageFont,ImageDraw
import re

class Xhs:
    def __init__(self):
        with open(os.path.join(os.path.dirname(__file__),'configs','poster.config'),'r',encoding='utf-8') as f:
            lines=f.readlines()
            _line=''
            for line in lines:
                newLine=line.strip('\n')
                _line=_line+newLine
            config=json.loads(_line)

        self.crs_info=config['科学机器人课程信息表']
        self.drawing_dir=config['图纸文件夹']
        self.font_list=config['fontList']
        self.template_bg_src=config['小红书模板图片']


    def read_crs(self):
        df_crs=pd.read_excel(self.crs_info)
        
        return df_crs

    def crs_poster_txt(self,crs_name='L178避障赛车'):
        df_crs=self.read_crs()
        all_crs_info=df_crs[df_crs['课程编号']==crs_name[:4]]
        knlg=all_crs_info['知识点'].tolist()[0]
        crs_title=crs_name[4:]

        return crs_title,knlg

    def draw_poster(self,crs_name='L178避障赛车'):
        txts=self.crs_poster_txt(crs_name=crs_name)
        bg=Image.open(self.template_bg_src)

        crs_pic=Image.open(os.path.join(self.drawing_dir,crs_name,crs_name[4:]+'.jpg'))
        crs_pic=crs_pic.crop((560,0,1280,720))
        crs_pic=crs_pic.resize((274,274))
        bg.paste(crs_pic,(221,164))

        draw=ImageDraw.Draw(bg)
        title_size=30
        x_title=bg.size[0]//2-len(txts[0])*title_size//2
        draw.text((x_title,470),txts[0],fill='#2CA6E0',font=composing.TxtFormat().fonts(font_name='方正韵动中黑简',font_size=title_size))
        composing.TxtFormat().put_txt_img(draw,tt=txts[1],total_dis=350,xy=(220,660),dis_line=18,fill='#E95513',font_name='楷体',font_size=24,addSPC='no')
        # print(bg.size)

        return bg

    def save_xhs_cover(self,crs_name='L178避障赛车',save_dir='e:/temp/temp_dzxc/小红书课程预告封面',open_tg='yes'):
        print('\n正在生成 {} 的小红书封图……'.format(crs_name),end='')
        pic=self.draw_poster(crs_name=crs_name)
        print('完成\n\n正在保存 {} 的小红书封图'.format(crs_name),end='')
        new_save_dir=os.path.join(save_dir,crs_name)
        if not os.path.exists(new_save_dir):
            os.makedirs(new_save_dir)
        save_name=os.path.join(new_save_dir,crs_name+'.jpg')    
        pic.save(save_name,quality=95,subsampling=0)
        print('完成')

        if open_tg=='yes':
            os.startfile(save_dir)

    def batch_deal_xhs(self,open_tg='no'):
        df_crs_list=self.read_crs()
        df_crs_list['id_crs']=df_crs_list['课程编号']+df_crs_list['课程名称']
        crs_list=[]
        for crs in df_crs_list['id_crs'].tolist():
            if re.match(r'^L\d\d\d.*',crs):
                crs_list.append(crs)
        
        for lego_crs in crs_list:
            try:
                self.save_xhs_cover(crs_name=lego_crs,open_tg=open_tg)
            except Exception as e:
                print(lego_crs,e)

        print('完成')

            


if __name__=='__main__':
    p=Xhs()
    # p.save_xhs_cover(crs_name='L028手动雨刮器',open_tg='no')
    p.batch_deal_xhs(open_tg='no')