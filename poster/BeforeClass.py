import os
import sys
sys.path.append('i:/py/dzxc/module')
from pics_modify import Shape
import composing
import readConfig
from PIL import Image,ImageDraw
import pandas as pd
import re
from datetime import datetime

class LegoClass:
    def __init__(self,place='001-超智幼儿园'):
        config=readConfig.readConfig(os.path.join(os.path.dirname(os.path.realpath(__file__)),'configs','poster.config'))
        self.bg=config['科学机器人课模板图片']
        self.crs_info=config['科学机器人课程信息表']
        self.df=pd.read_excel(self.crs_info,sheet_name='课程信息')
        self.crs_proj_dir=config['图纸文件夹']
        self.font_list=config['fontList']
        self.font=composing.TxtFormat().fonts
        self.out_dir=config['输出文件夹'].replace('$',place)

    def read_excels(self,crs_name_input):        
        df_crs=self.df[self.df['课程编号']==crs_name_input[0:4]]
        knlg=df_crs['知识点'].tolist()[0]
        knlg=re.sub(r'\d.','· ',knlg)
        crs_name=df_crs['课程名称'].tolist()[0]
        
        return [crs_name,knlg]

    def exp_poster(self,date_crs_input='20210306',time_crs_input='1100-1230',addr_crs='超智幼儿园  五楼',crs_name_input='L063汽车雨刮器'):
        txts=self.read_excels(crs_name_input)
        weekday=datetime.strptime(str(date_crs_input),'%Y%m%d').weekday()+1 #通日期计算星期
        wd={'1':'一','2':'二','3':'三','4':'四','5':'五','6':'六','7':'日',}
        date_crs=date_crs_input[0:4]+'年'+date_crs_input[4:6]+'月'+date_crs_input[6:]+'日（星期'+wd[str(weekday)]+'）'
        time_crs=time_crs_input[0:2]+':'+time_crs_input[2:4]+' - '+time_crs_input[5:7]+':'+time_crs_input[7:]
        crs_name=txts[0] #课程名称
        knlg=txts[1]# 知识点

        bg=Image.open(self.bg)
        draw=ImageDraw.Draw(bg)
        ft_size_crs_name=55
        x_crs_name=int(bg.size[0]-len(crs_name)*ft_size_crs_name)/2 #居中
        draw.text((x_crs_name,351),crs_name,fill='#00a0e9',font=self.font('优设标题',ft_size_crs_name)) #课程名称
        draw.text((260,178),date_crs,fill='#595757',font=self.font('上首金牛体',34)) #课程日期
        draw.text((260,228),time_crs,fill='#595757',font=self.font('上首金牛体',34)) #课程日期
        draw.text((260,279),addr_crs,fill='#595757',font=self.font('上首金牛体',34)) #课程地点
        composing.TxtFormat().put_txt_img(draw=draw,tt=knlg,total_dis=480,xy=[120,530],dis_line=20, \
                                fill='#595757',font_name='丁永康硬笔楷书',font_size=34,addSPC='None') #知识点

        _img=Image.open(os.path.join(self.crs_proj_dir,crs_name_input,crs_name_input[4:]+'.jpg'))
        img=_img.resize((int(300*_img.size[0]/_img.size[1]),300))
        w,h=img.size
        new_img=img.crop((int(w*5/16),0,w,h))
        new_img_round=Shape().circle_corner(new_img,radii=150)
        # new_img_round.show()
        bg.paste(new_img_round,(int((bg.size[0]-new_img_round.size[0])/2),700),mask=new_img_round)
        jpg=bg.convert('RGB')
        save_name=os.path.join(self.out_dir,date_crs_input+'_'+crs_name_input+'_课前海报.jpg')
        jpg.save(save_name,quality=85)
        os.startfile(self.out_dir)
        # bg.show()
        print('完成')

if __name__=='__main__':
    my=LegoClass()
    my.exp_poster(date_crs_input='20210306',time_crs_input='1100-1230',crs_name_input='L063汽车雨刮器')
    # my.read_excels()