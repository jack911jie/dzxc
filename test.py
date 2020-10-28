import os
import pandas as pd
import re

dirName='i:\\每周乐高课'

df=pd.read_excel('i:\\大智小超\\课程表\\课程信息表.xlsx')

dirs=['20200922L037认识零件','20200929L033双翼飞机','20201013L034夏天的手摇风扇','20201020L035啃骨头的小狗','20201024L037认识零件']

# for dir in dirs:
#     oldname=os.path.join(dirName,dir)
#     crsName=dir[8:]
#     # print(crsName)
#     crsCode=df[df['课程名称']==crsName]['课程编号'].values.tolist()[0]
#     # print(crsCode)
#     newname=os.path.join(dirName,dir[:8]+crsCode+dir[8:])
#     print(newname)
#     os.rename(oldname,newname)

ptn='\d{8}-.*-\d{3}.JPG'

for dir in dirs:
    for file in os.listdir(os.path.join(dirName,dir)):
        if re.match(ptn,file):
            crsName=file.split('-')[1]
            crsCode=df[df['课程名称']==crsName]['课程编号'].values.tolist()[0]
            oldname=os.path.join(dirName,dir,file)
            newname=os.path.join(dirName,dir,file[:9]+crsCode+file[9:])
            os.rename(oldname,newname)