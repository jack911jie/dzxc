import os
import numpy as np
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

if __name__=='__main__':
    pic=draw()
    pic.radar([3,4,3,4,5,5,2])