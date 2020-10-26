import re
from PIL import Image,ImageFont,ImageDraw
import logging
logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(funcName)s-%(lineno)d - %(message)s')
logger = logging.getLogger(__name__)

def char_len(txt):
    len_s=len(txt)
    len_u=len(txt.encode('utf-8'))
    ziShu_z=(len_u-len_s)/2
    ziShu_e=len_s-ziShu_z
    total=ziShu_z+ziShu_e*0.5
    return total

def fonts(font_name,font_size):
    fontList={
        '腾祥金砖黑简':'c:\\windows\\fonts\\腾祥金砖黑简.TTF',
        '汉仪糯米团':'j:\\fonts\\HYNuoMiTuanW.ttf',
        '丁永康硬笔楷书':'j:\\fonts\\2012DingYongKangYingBiKaiShuXinBan-2.ttf',
        '微软雅黑':'i:\\py\\msyh.ttf',
        '鸿蒙印品':'j:\\fonts\\hongMengHei.ttf',
        '优设标题':'j:\\fonts\\yousheTitleHei.ttf',
        '汉仪超级战甲':'j:\\fonts\\HYChaoJiZhanJiaW-2.ttf',
        '汉仪心海行楷w':'j:\\fonts\\HYXinHaiXingKaiW.ttf',
        '华康海报体W12(p)':'j:\\fonts\\HuaKangHaiBaoTiW12-P-1.ttf',
        '汉仪锐智w':'j:\\fonts\\HYRuiZhiW.ttf',
        '杨任东竹石体':'j:\\fonts\\yangrendongzhushi-Regular.ttf',
        '楷体':'C:\\Windows\\Fonts\\simkai.ttf',
        '汉仪字酷堂义山楷w':'j:\\fonts\\HYZiKuTangYiShanKaiW-2.ttf'                    
    }

    # ImageFont.truetype('j:\\fonts\\2012DingYongKangYingBiKaiShuXinBan-2.ttf',font_size)


    return ImageFont.truetype(fontList[font_name],font_size)

def split_txt_Chn_eng(wid,font_size,txt_input,Indent='no'):
    
    def put_words(txts,zi_per_line):    
        txtGrp=[]
        wd_lng=0
        pre_txt=''
        for c,t in enumerate(txts):
            if res.match(t):
                if wd_lng+1>zi_per_line:
                    txtGrp.append(pre_txt)
                    pre_txt=t
                    wd_lng=1
                else:
                    pre_txt=pre_txt+t
                    wd_lng=wd_lng+1
            else:
                if wd_lng+char_len(t)>zi_per_line: #先判断是这个英文单词+原有拼接的字符串长度是否>每行字符数
                    txtGrp.append(pre_txt) #大于，则保持原有的拼接字符串，不再加入该英文单词
                    pre_txt=' '+t #新的英文单词另起一行
                    wd_lng=char_len(t) #拼接字符串长度清零重计
                else:                    
                    pre_txt=pre_txt+' '+t
                    wd_lng=wd_lng+char_len(t)
            
            if wd_lng>zi_per_line:
                wd_lng=0                
                txtGrp.append(pre_txt)
                pre_txt=''                
            else:
                if c==len(txts)-1:
                    wd_lng=0                
                    txtGrp.append(pre_txt)
                    pre_txt='' 
                    
                
        logging.info(txtGrp)
        return txtGrp
    
    txts=txt_input.splitlines()
    logging.info(txts)    
    
    if Indent=='yes':
        for i,t in enumerate(txts):
            txts[i]=chr(12288)+chr(12288)+t
    
    res = re.compile(r'([\u4e00-\u9fa5])')   
    singleTxts=[]
    for t in txts:
        _t=res.split(t)
        _t_no_empty=list(filter(None,_t))      
        singleTxts.append(_t_no_empty)        
    
    split_txt=[]
    for singleTxt in singleTxts:
        grp=[]
        for st in singleTxt:
            if res.match(st):
                grp.append(st)
            else:
                grp.extend(st.split(' '))
        split_txt.append(grp)
        
    logging.info(split_txt)
        
    
    total_num=0
    for r in split_txt:
        total_num=total_num+len(r)
        
    zi_per_line=int(wid//font_size)
           
    outTxt=[]
    for r,split_t in enumerate(split_txt):
        outTxt.append(put_words(split_t,zi_per_line))
        
    
   
    return outTxt 

def put_txt_img(draw,t,total_dis,xy,dis_line,fill,font_name,font_size,addSPC='None'):
        
    fontInput=fonts(font_name,font_size)            
    if addSPC=='add_2spaces': 
        Indent='yes'
    else:
        Indent='no'
        
    # txt=self.split_txt(total_dis,font_size,t,Indent='no')
    txt=split_txt_Chn_eng(total_dis,font_size,t,Indent='no')
    # font_sig = self.fonts('丁永康硬笔楷书',40)
    num=len(txt)   
    # draw=ImageDraw.Draw(img)

    logging.info(txt)
    n=0
    for t in txt:              
        m=0
        for tt in t:                  
            x,y=xy[0],xy[1]+(font_size+dis_line)*n
            if addSPC=='add_2spaces':   #首行缩进
                if m==0:    
                    # tt='  '+tt #首先前面加上两个空格
                    logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                    logging.info(tt)
                    draw.text((x+font_size*0.2,y), tt, fill = fill,font=fontInput) 
                else:                       
                    logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                    logging.info(tt)
                    draw.text((x,y), tt, fill = fill,font=fontInput)  
            else:
                # logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                # logging.info(tt)
                draw.text((x,y), tt, fill = fill,font=fontInput)  

            m+=1
            n+=1

def char_len(txt):
    len_s=len(txt)
    len_u=len(txt.encode('utf-8'))
    ziShu_z=(len_u-len_s)/2
    ziShu_e=len_s-ziShu_z
    total=ziShu_z+ziShu_e*0.5    
    return total