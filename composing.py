import re
import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(funcName)s-%(lineno)d - %(message)s')
logger = logging.getLogger(__name__)

def char_len(txt):
    len_s=len(txt)
    len_u=len(txt.encode('utf-8'))
    ziShu_z=(len_u-len_s)/2
    ziShu_e=len_s-ziShu_z
    total=ziShu_z+ziShu_e*0.5
    return total

def split_txt_Chn_eng(wid,font_size,txt_input,Indent='no'):
    
    def put_words(txts,zi_per_line):    

        txtGrp=[]
        wd_lng=0
        pre_txt=''
        for c,t in enumerate(txts):
            if res.match(t):
                pre_txt=pre_txt+t #汉字
                wd_lng=wd_lng  #汉字，字数按加1
            else:
                pre_txt=pre_txt+' '+t #英文前加个空格
                wd_lng=wd_lng+char_len(t) #英文字数按英文字符数计算
            
            if wd_lng>zi_per_line:
                wd_lng=0                
                txtGrp.append(pre_txt)
                pre_txt=''                
            else:
                if c==len(txts)-1:
                    wd_lng=0                
                    txtGrp.append(pre_txt)
                    pre_txt=''         
        return txtGrp
    
    txts=txt_input.splitlines()
    logging.info(txts)    
    
    if Indent=='yes':
        for i,t in enumerate(txts):
            txts[i]=chr(12288)+chr(12288)+t
    
    res = re.compile(r'([\u4e00-\u9fa5])')   #检测汉字
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
        
#    logging.info(split_txt)
        
    
    total_num=0
    for r in split_txt:
        total_num=total_num+len(r)
        
    zi_per_line=int(wid//font_size)
           
    outTxt=[]
    for r,split_t in enumerate(split_txt):
        outTxt.append(put_words(split_t,zi_per_line))
   
    return outTxt 