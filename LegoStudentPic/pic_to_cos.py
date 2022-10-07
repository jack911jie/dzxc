import os
import sys

import re
import shutil
import pandas as pd
from datetime import datetime
from urllib.parse import quote

class Picutres:
    def __init__(self):
        pass

    def get_list_copy(self,dir='E:\\大智小超\\课后照片及反馈',tgt_dir='E:\\大智小超\\上传COS'):
        # for directs in os.listdir(dir):
        #     if os.path.isdir(os.path.join(dir,directs)) and re.match(r'\d{8}-L\d{3}.*',directs):
        #         print(directs)
        print('正在复制文件')
        for root,dirs,files in os.walk(dir):
            for fn in files:

                if re.match(r'\d{8}-L\d{3}.*',fn):
                    cplt_name=os.path.join(root,fn)
                    split_name=cplt_name.split('\\')
                    # print(split_name)
                    std_name=split_name[2]
                    yr=fn[:4]
                    pre_name=std_name+'-'+split_name[-1]
                    save_dir=os.path.join(tgt_dir,yr)
                    if not os.path.exists(save_dir):
                        os.makedirs(save_dir)
                    save_name=os.path.join(save_dir,pre_name)
                    
                    print(cplt_name,'--->',save_name)

                    shutil.copyfile(cplt_name,save_name)
        print('完成')

    def to_csv(self,dir='E:\\大智小超\\上传COS'):
        _df=[]
        for root,dirnames,files in os.walk(dir):
            for fn in files:
                if re.match(r'.*-\d{8}-L\d{3}.*',fn):
                    _df.append(fn)
        
        # print(_df)
        df=pd.DataFrame({'filename':_df})

        df.to_csv('e:/temp/temp_dzxc/pic_info.csv')
        print('完成')

    def make_table(self):
        df=pd.read_csv('e:/temp/temp_dzxc/pic_info.csv')
        df['图片上传日期']=df['filename'].apply(lambda x: datetime.strptime(x.split('-')[1],'%Y%m%d'))
        df['学生姓名']=df['filename'].apply(lambda x: re.findall(r'[\u4e00-\u9fa5]{2,}',x.split('-')[0])[0])
        df['所属课程']='乐高'
        df['课程编号']=df['filename'].apply(lambda x: x.split('.')[0].split('-')[1]+'-'+x.split('.')[0].split('-')[2])
        df['课程名称']=df['filename'].apply(lambda x:x.split('-')[2])
        df['图片名']=df['filename'].apply(lambda x:x.split('-')[2][4:]+'-'+x.split('-')[3][:-4])
        df['图片类型']='课堂照片'
        df['排序']=df['图片名'].apply(lambda x:int(x.split('-')[1]))
        df['ImgURL']=df['filename'].apply(lambda x: self.url(x.split('-')[1][:4]+'/'+x,prefix='https://chuntianhuahua-1257410889.cos.ap-guangzhou.myqcloud.com/dzxc/001-CHAOZHI/'))

        df_out=df[['图片上传日期','学生姓名','所属课程','课程编号','课程名称','图片名','图片类型','排序','ImgURL']]

        return df_out

    def url(self,url_input='2021/CJY陈锦媛-20210924-L107我的小房子-017.JPG',prefix='https://chuntianhuahua-1257410889.cos.ap-guangzhou.myqcloud.com/dzxc/001-CHAOZHI/'):
        url_out=quote(os.path.join(prefix,url_input))
        url_out=url_out.replace('%3A//','://')

        # print(url_out)

        return url_out

        
if __name__=='__main__':
    p=Picutres()
    # p.get_list_copy(dir='E:/temp/temp_dzxc/temp0429')
    # p.to_csv()
    res=p.make_table()
    res.to_excel('e:/temp/temp_dzxc/学生图片库表.xlsx')
    # url=os.path.join('2021','CJY陈锦媛-20210924-L107我的小房子-017.JPG')
    # res=p.url('2021/CJY陈锦媛-20210924-L107我的小房子-017.JPG')
    # print(res)