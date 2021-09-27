import os
from PIL import Image,ImageDraw

def circle_corner(img,radii=150):
    radii=int(img.size[0]*radii/4032)
    # 画圆（用于分离4个角）  
    circle = Image.new('L', (radii * 2, radii * 2), 0)  # 创建黑色方形
    # circle.save('1.jpg','JPEG',qulity=100)
    draw = ImageDraw.Draw(circle)
    draw.ellipse((0, 0, radii * 2, radii * 2), fill=255)  # 黑色方形内切白色圆形
    # circle.save('2.jpg','JPEG',qulity=100)

    # 原图转为带有alpha通道（表示透明程度）
    img = img.convert("RGBA")
    w, h = img.size

    # 画4个角（将整圆分离为4个部分）
    alpha = Image.new('L', img.size, 255)	#与img同大小的白色矩形，L 表示黑白图
    # alpha.save('3.jpg','JPEG',qulity=100)
    alpha.paste(circle.crop((0, 0, radii, radii)), (0, 0))  # 左上角
    alpha.paste(circle.crop((radii, 0, radii * 2, radii)), (w - radii, 0))  # 右上角
    alpha.paste(circle.crop((radii, radii, radii * 2, radii * 2)), (w - radii, h - radii))  # 右下角
    alpha.paste(circle.crop((0, radii, radii, radii * 2)), (0, h - radii))  # 左下角
    # alpha.save('4.jpg','JPEG',qulity=100)

    img.putalpha(alpha)		# 白色区域透明可见，黑色区域不可见
    # img.save('e:\\temp\\5555.png','PNG',qulity=100)

    # img.show()
    # img=img.convert('RGBA')
    return img

def round_pic(img,method='in'):
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