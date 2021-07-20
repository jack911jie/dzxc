import numpy as np
import matplotlib.pyplot as plt 
import matplotlib.font_manager as fm
from matplotlib.backends.backend_agg import FigureCanvasAgg
from PIL import Image,ImageDraw
import base64

def mat_to_pil_img(fig):
    # 将plt转化为numpy数据
    canvas = FigureCanvasAgg(plt.gcf())
    # 绘制图像
    canvas.draw()
    # 获取图像尺寸
    w, h = canvas.get_width_height()
    # 解码string 得到argb图像
    buf = np.frombuffer(canvas.tostring_argb(), dtype=np.uint8)
    fig.canvas.draw()
    # 获取图像尺寸
    w, h = fig.canvas.get_width_height()
    # 获取 argb 图像
    buf = np.frombuffer(fig.canvas.tostring_argb(), dtype=np.uint8)
    # 重构成w h 4(argb)图像
    buf.shape = (w, h, 4)
    # 转换为 RGBA
    buf = np.roll(buf, 3, axis=2)
    # 得到 Image RGBA图像对象 (需要Image对象的同学到此为止就可以了)
    image = Image.frombytes("RGBA", (w, h), buf.tobytes())
    # # 转换为numpy array rgba四通道数组
    # image = np.asarray(image)
    # # 转换为rgb图像
    # rgb_image = image[:, :, :3]
    return image

def round_corner(img,method='in'):
    if img.mode!='RGBA':
        img=img.convert('RGBA')

    if method=='out':
        w,h=img.size
        r=int(np.sqrt(w*w/4+h*h/4))
        d=int(r*2)
        bg=Image.new('RGBA',(d,d),'#FFFFFF')
        scale=3
        alpha_layer = Image.new('L', (d*scale, d*scale), 0)
        draw=ImageDraw.Draw(alpha_layer)
        draw.ellipse((0,0,d*scale,d*scale),fill='#FFFFFF')
        alpha_layer=alpha_layer.resize((d,d))
        
        # bg.paste(alpha_layer,(0,0),mask=alpha_layer)
        bg.paste(img,((d-w)//2,(d-h)//2),mask=img)
        bg.putalpha(alpha_layer)
        # bg.save('e:/temp/dq.png')
    elif method=='in':        
        w,h=img.size
        d=min(w,h)
        scale=3
        alpha_layer=Image.new('L',(w*scale,h*scale),0)
        draw=ImageDraw.Draw(alpha_layer)
        if h>=w:
            draw.ellipse((0,(h*scale-d*scale)//2,w*scale,(h*scale-d*scale)//2+d*scale),fill='#FFFFFF')
            alpha_layer=alpha_layer.resize((w,h))
            img.putalpha(alpha_layer)
        else:
            draw.ellipse(((w*scale-d*scale)//2,0,(w*scale-d*scale)//2+d*scale,h*scale),fill='#FFFFFF')
            alpha_layer=alpha_layer.resize((w,h))
            img.putalpha(alpha_layer)
        bg=img

        # alpha_layer.show()
        # img.show()
        # bg.show()
    return bg

def mat_test():
    x=[2,4,6,8]
    y=[1,2,3,4]
    fig=plt.figure(figsize=(9,20))
    ax1=fig.add_axes([0.1, 0.08, 0.8, 0.12],facecolor='#FCF8FF')
    ax1.plot(x,y)
    # plt.show()

    return fig

def pic_to_base64(pic):
    f=open(pic,'rb') #二进制方式打开图文件
    ls_f=base64.b64encode(f.read()) #读取文件内容，转换为base64编码
    f.close()    
    return ls_f

if __name__=='__main__':
    # img=Image.open('E:\\健身项目\\素材\\男性头像01.jpg')
    # img=Image.open('E:\\铭湖健身\\素材\\女性头像02.jpg')
    # imgg=round_corner(img,method='in')
    # imgg.save('E:\\铭湖健身\\素材\\tt.png')
    # imgg.show()
    k=mat_test()
    p=mat_to_pil_img(k)