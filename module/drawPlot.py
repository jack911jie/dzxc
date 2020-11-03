import os
import random
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import patches
plt.rcParams['font.sans-serif']=['SimHei']#中文出现需  u'内容'
plt.rcParams['axes.unicode_minus']=False

class draw:
    def __init__(self):
        pass

    def radar(self,data):
        mypath = 'i:\\py\\dzxc'
        savefilename = os.path.join(mypath,'testRadar.jpg')
        #print('雷达图文件名：', savefilename)

        ljl,kjxxl,ljsw,zyl,czl,bdl,kcnl=data
        # 构造数据
        values = [ljl,kjxxl,ljsw,zyl,czl,bdl,kcnl]
        feature = ['理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力']

        N = len(values)
        # 设置雷达图的角度，用于平分切开一个圆面
        angles = np.linspace(0, 2 * np.pi, N, endpoint=False)


        # 为了使雷达图一圈封闭起来，需要下面的步骤
        values = np.concatenate((values, [values[0]]))
        angles = np.concatenate((angles, [angles[0]]))

        print(values,angles)

        # 绘图
        fig = plt.figure(figsize=(6.8,4.8))
        # 这里一定要设置为极坐标格式
        ax = fig.add_subplot(111, polar=True)
        # ccl=ax.patch

        # 绘制折线图
        ax.plot(angles, values, 'o-', linewidth=3,color='#218FBD')
        # 填充颜色
        ax.fill(angles, values, alpha=0.25)
        # 添加每个特征的标签
        ax.set_thetagrids(angles * 180 / np.pi, '',color='r',fontsize=13)
        # 设置雷达图的范围
        r_distance=6
        ax.set_rlim(0, r_distance)

        ax.grid(color='g', alpha=0.25, lw=3)
        ax.spines['polar'].set_color('grey')
        ax.spines['polar'].set_alpha(0.2)
        ax.spines['polar'].set_linewidth(3)
        # ax.spines['polar'].set_linestyle('-.')

        #项目名称：
        a=[0,0,np.pi/30,-np.pi/50,0,0,0]
        b=[r_distance*1.1,r_distance*1.1,r_distance*1.2,r_distance*1.4,r_distance*1.5,r_distance*1.2,r_distance*1.1]

        feature2 = ['理解力','空间想象力','逻辑思维','注意力','创造力','表达力','抗挫能力']
        for k,i in enumerate(angles):
            try:
                print(k,i,feature2[k])
                ax.text(i+a[k],b[k],feature2[k],fontsize=18,color='#565656')
            except:
                pass

        #分值：
        c = [1, 0.6, 1.6, 2.3, 1.5, 1,1]
        print(len(angles))
        for j,i in enumerate(angles):
            try:
                r=values[j]-2*i/np.pi
                ax.text(i,values[j]+c[j],values[j],color='#218FBD',fontsize=18)
            except:
                pass

        # 添加标题
        #plt.title('活动前后员工状态表现')
        # 添加网格线
        ax.grid(True,color='grey',alpha=0.1)

        # a=np.arange(0,2*np.pi,0.01)
        # ax.plot(a,10*np.ones_like(a),linewidth=2,color='b')


        ax.set_yticklabels([])
        plt.savefig(savefilename,transparent=True,bbox_inches='tight')
        # 显示图形
        plt.show()

    def rose(self):
        #读取数据
        data=pd.read_excel('e:\\temp\\海外疫情.xlsx',index_col=0)

        #数据计算，这里只取前20个国家
        radius = data['累计'][:20]
        n=radius.count()
        theta = np.arange(0, 2*np.pi, 2*np.pi/n)+2*np.pi/(2*n)    #360度分成20分，外加偏移

        #在画图时用到的 plt.cm.spring_r(r)   r的范围要求时[0,1]
        radius_maxmin=(radius-radius.min())/(radius.max()-radius.min())  #x-min/max-min   归一化  

        #画图
        fig = plt.figure(figsize=(20,5),dpi=256)
        ax = fig.add_subplot(projection='polar')    #启用极坐标
        bar = ax.bar(theta, radius,width=2*np.pi/n)


        ax.set_theta_zero_location('N')  #分别为N, NW, W, SW, S, SE, E, NE
        ax.set_rgrids([])    #用于设置极径网格线显示
        # ax.set_rticks()    #用于设置极径网格线的显示范围
        # ax.set_theta_direction(-1)    #设置极坐标的正方向
        ax.set_thetagrids([])  #用于设置极坐标角度网格线显示
        # ax.set_theta_offset(np.pi/2)       #用于设置角度偏离
        ax.set_title('新冠肺炎全球疫情形势',fontdict={'fontsize':8})   #设置标题

        #设置扇形各片的颜色
        for r, bar in zip(radius_maxmin, bar):
            bar.set_facecolor(plt.cm.spring_r(r))  
            bar.set_alpha(0.8)

        #设置边框显示    
        for key, spine in ax.spines.items():  
            if key=='polar':
                spine.set_visible(False)

        plt.show()

        #保存图片
        fig.savefig('COVID.png')

    def rose2(self):
        # 参考https://www.sohu.com/a/223885806_718302
        # y=20
        # x=np.pi/2
        # w=np.pi/2

        # color=(206/255,32/255,69/255)
        # edgecolor=(206/255,32/255,69/255)

        # fig=plt.figure(figsize=(13.44,7.5))#建立一个画布
        # ax=fig.add_subplot(111,projection='polar')#建立一个坐标系，projection='polar'表示极坐标
        # ax.bar(x=x,height=y,width=w,bottom=10,color=color,edgecolor=color)
        # # fig.savefig('E:test.png',dpi=400,bbox_inches='tight',transparent=True)

        x1=[np.pi/10+np.pi*i/5 for i in range(1,11)]
        x2=[np.pi/20+np.pi*i/5 for i in range(1,11)]
        x3=[3*np.pi/20+np.pi*i/5 for i in range(1,11)]
        y1=[7000 for i in range(0,10)]
        y2=[6000 for i in range(0,10)]
        
        fig=plt.figure(figsize=(13.44,7.5))
        ax = fig.add_subplot(111,projection='polar')
        ax.axis('off')
        ax.bar(x=x1, height=y1,width=np.pi/5,color=(220/255,222/255,221/255),edgecolor=(204/255,206/255,205/255))
        ax.bar(x=x1, height=y2,width=np.pi/5,color='w',edgecolor=(204/255,206/255,205/255))

        random.seed(100)
        y4=[random.randint(4000,5500) for i in range(10)]
        y5=[random.randint(3000,5000) for i in range(10)]
        
        ax.bar(x=x2, height=y4,width=np.pi/10,color=(206/255,32/255,69/255),edgecolor=(206/255,32/255,69/255))
        ax.bar(x=x3, height=y5,width=np.pi/10,color=(34/255,66/255,123/255),edgecolor=(34/255,66/255,123/255))

        y6=[2000 for i in range(0,10)]
        ax.bar(x=x1, height=y6,width=np.pi/5,color='w',edgecolor='w')

        plt.show()




if __name__=='__main__':
    pic=draw()
    # pic.radar([3,4,3,4,5,5,2])
    pic.rose2()