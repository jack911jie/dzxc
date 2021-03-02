import os
import sys
import matplotlib.pyplot as plt
import numpy as np

class dat_summary:
    def __init__(self):
        pass

    def rose(self):
        dat=['理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力','协作能力']
        score=[3,3,4,5,2,2,3,5]
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
                width = 0.3,
                color = np.random.random((len(score),3)),
                # labels=str(country_list), 
                align = 'edge')
        ''' 
            显示一些简单的中文图例
        '''
        plt.rcParams['font.sans-serif']=['SimHei']  # 黑体
        ax.set_title('能力测评',fontdict={'fontsize':20})
        for angle,scores,data in zip(theta,score,dat):
            ax.text(angle+0.04,scores+0.2,data) 


        plt.axis('off')

        # plt.savefig('Nightingale_rose.png')
        plt.show()


        print(len(score))


if __name__=='__main__':
    my=dat_summary()
    my.rose()