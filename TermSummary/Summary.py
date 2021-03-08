import os
import sys
sys.path.append('i:/py/dzxc/module')
import composing
import WashData
import days_calculate
from readConfig import readConfig
import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np

class dat_summary:
    def __init__(self):
        config=readConfig(os.path.join(os.path.dirname(os.path.realpath(__file__)),'configs','term_summary_config.dazhi'))
        self.std_dir=config['学生信息文件夹']
        
    def draw_std_term_crs(self,std_name='韦宇浠',start_date='20200930',end_date='20210121',weekday='2'):
        std_name=std_name.strip()
        xls_name=os.path.join(self.std_dir,'2020乐高课程签到表（周'+days_calculate.num_to_ch(weekday)+'）.xlsx')
        std_term_crs=WashData.std_term_crs(std_name=std_name,start_date=start_date,end_date=end_date,xls=xls_name)
        print(std_term_crs)

    def rose(self,std_name='韦宇浠',weekday=2):
        df_ability=WashData.std_feedback_ability(os.path.join(self.std_dir,'每周课程反馈','学员课堂学习情况反馈表.xlsx'),weekday=weekday)
        std_ability=df_ability[df_ability['姓名']==std_name].iloc[:,1:]
        dat=std_ability.columns.tolist()
        score=std_ability.iloc[0,:].tolist()

        # dat=['理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力','协作能力']
        # score=[3,3,4,5,2,2,3,5]
        score.sort()
        theta = np.linspace(0,2*np.pi,len(dat),endpoint=False)    # 360度等分成n份

        # 设置画布
        fig = plt.figure(figsize=(12,10))
        # 极坐标
        ax = plt.subplot(111,projection = 'polar')
        # 顺时针并设置N方向为0度
        ax.set_theta_direction(-1)
        ax.set_theta_zero_location('N') 

        # 在极坐标中画柱形图
        ax.bar(theta,
                score,
                width = 0.75,
                color = np.random.random((len(score),3)),
                # labels=str(country_list), 
                align = 'edge')
        ''' 
            显示一些简单的中文图例
        '''
        plt.rcParams['font.sans-serif']=['SimHei']  # 黑体
        title=std_name+'同学能力测评'
        ax.set_title(title,fontdict={'fontsize':20,'color':'#3923a8'})
        for angle,scores,data in zip(theta,score,dat):
            ax.text(angle+0.1,scores+0.2,data) 
            # ax.text(angle+0.04,scores+0.6,str(scores))

        ax.axis('off')
        # plt.savefig('Nightingale_rose.png')
        plt.show()


        # print(len(score))


if __name__=='__main__':
    my=dat_summary()
    # my.rose(std_name='陶盛挺',weekday=2)
    my.draw_std_term_crs(std_name='韦宇浠',start_date='20200922',end_date='20201201',weekday='2')