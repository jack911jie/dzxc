import os
import sys
import fitz

class JpgImport:
    def __init__(self):
        pass

    def dec(start): #接收body函数，不用self
        print(start,'……',end='')
        def print_wrap(func):#接收body函数，不用self
            
            def wrap(self):#需要用self，接收的是body的类实例
                # print(end)
                return func(self)
                
            return wrap
        
        
        return print_wrap

    @dec(start='开始读取pdf')
    def read_jpg(self,pdf_fn='e:/temp/tttttt.pdf'):
        pdf_doc = fitz.open(pdf_fn)
        for pg in range(pdf_doc.pageCount):
            page = pdf_doc[pg]
            rotate = int(0)
            # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
            # 此处若是不做设置，默认图片大小为：792X612, dpi=72
            # zoom_x = 1.33333333 #(1.33333333-->1056x816)   (2-->1584x1224)
            # zoom_y = 1.33333333
            zoom_x=4.1666666666666 #(默认宽792———>2481)
            zoom_y=4.1666666666666 #(默认高612———>3508)
            mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
            pix = page.getPixmap(matrix=mat, alpha=False)
            
            # if not os.path.exists(imagePath):#判断存放图片的文件夹是否存在
            #     os.makedirs(imagePath) # 若图片文件夹不存在就创建
        
            pix.writePNG('e:/temp/pdf_jpg/'+str(pg)+'.jpg')#将图片写入指定的文件夹内

        print('完成')

    

if __name__=='__main__':
    p=JpgImport()
    p.read_jpg()