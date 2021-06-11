from PIL import Image,ImageEnhance
from pyzbar import pyzbar
import qrcode
import iptcinfo3

pic='e:\\temp\\qrcodetest03.jpg'
pic2='I:\\每周乐高课_学员\\ZWY周琬瑜\\20200929-L033双翼飞机-008.JPG'
pic3='e:\\temp\\双翼机.JPG'

def make_qrcode():
    qr=qrcode.QRCode(version = 2,error_correction = qrcode.constants.ERROR_CORRECT_L,box_size=10,border=10,)
    qr.add_data(['H001','揭东豆'])
    qr.make(fit=True)
    img = qr.make_image()
    img.show()
    #img.save('D:/test.jpg')

def detect_and_read_qrcode():

    img = Image.open(pic)
    # img = ImageEnhance.Brightness(img).enhance(2.0)#增加亮度
 
    # img = ImageEnhance.Sharpness(img).enhance(17.0)#锐利化
 
    # img = ImageEnhance.Contrast(img).enhance(4.0)#增加对比度
 
    # img = img.convert('L')#灰度化
    texts = pyzbar.decode(img)

    for text in texts:
        data=text.data.decode('utf-8')
        print(data)

def write_jpgMark():
    img=iptcinfo3.IPTCInfo(pic3)
    img['keywords']=['H001','揭东豆']
    img.save()

write_jpgMark()