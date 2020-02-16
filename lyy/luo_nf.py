#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import openpyxl as opx
from openpyxl import Workbook
from openpyxl.styles import Alignment,Font,Side,Border,NamedStyle
from openpyxl.utils import get_column_letter
import xlrd
import os
import time
import socket

class sht_merge:
    def __init__(self,filename):
        self.filename=filename
    
    def rd(self):
        print("* 正在读取数据... \n")
        sht_0=pd.read_excel(self.filename,sheet_name=0)
        sht_1=pd.read_excel(self.filename,sheet_name=1,usecols=[0,1,2,3,4,5,6,7,8])
        sht_2=self.slct_ValidAgent() #先整理出有效的律所列表
        sht_3=self.slct_dyw() #先整理出有效的不动产列表
        
        print("完成 \n * 正在处理数据... \n")

        res_1=sht_0.sort_values(by="案件名称") #表一按名称排序
        cls_2=sht_1.loc[sht_1['审级']==2]  #二审
        cls_3=sht_1.loc[sht_1['审级']==3]  #再审
        # cls_4=sht_1.loc[sht_1['审级']==4]  #执行

        
        res=pd.merge(res_1,cls_2,left_on='案号',right_on='一审案号',how="left",suffixes=["","_二审"])                     .drop(columns=["一审案号","案件名称_二审","审级","主办_二审","副办_二审"])
        res=pd.merge(res,cls_3,left_on='案号',right_on='一审案号',how="left",suffixes=["","_再审"])                     .drop(columns=["一审案号","案件名称_再审","审级","审理情况_再审","主办法官_再审","管辖法院_再审","主办_再审","副办_再审"])  
        res=pd.merge(res,sht_2,left_on='案件名称',right_on='一审案件名称',how="left")                     .drop(columns=["一审案件名称"])
        res=pd.merge(res,sht_3,left_on='案件名称',right_on='案件名称',how="left")

        
        order=["户数","案件数","案件名称","起诉金额（元）","财务代偿余额（元）","抵押物情况","主办","副办","诉讼状态"                 ,"代理方式","律所名称","辅助机构"                 ,"案号","案号_二审","案号_再审","执行案号","起诉日期"                 ,"立案日期","一审情况","执行情况","一审管辖法院","一审主办法官"                 ,"审理情况","管辖法院","主办法官"                 ,"执行情况","执行法院","执行主办法官"                 ,"备注"]
        
        res=res[order]
        res=res.rename(columns={"案号":"一审案号","案号_二审":"二审案号","案号_再审":"再审案号"                            ,"审理情况":"二审审理情况","管辖情况":"二审管辖情况","主办法官":"二审主办法官"                            })

#         print(res)
        self.res=res
        print("完成 \n ")
    
    #处理代理机构表，挑选出有效的律所
    def slct_ValidAgent(self):

        sht_2=pd.read_excel(self.filename,sheet_name=2)

        dl_1=sht_2.loc[sht_2["一审代理律所名称"]==sht_2["目前有效"]]             .drop(columns=["二审代理律所名称","再审代理律所名称","执行代理律所名称"])             .rename(columns={"一审代理律所名称":"律所名称"})

        dl_2=sht_2.loc[sht_2["二审代理律所名称"]==sht_2["目前有效"]]             .drop(columns=["一审代理律所名称","再审代理律所名称","执行代理律所名称"])             .rename(columns={"二审代理律所名称":"律所名称"})

        dl_3=sht_2.loc[sht_2["再审代理律所名称"]==sht_2["目前有效"]]             .drop(columns=["一审代理律所名称","二审代理律所名称","执行代理律所名称"])             .rename(columns={"再审代理律所名称":"律所名称"})

        dl_4=sht_2.loc[sht_2["执行代理律所名称"]==sht_2["目前有效"]]             .drop(columns=["一审代理律所名称","二审代理律所名称","再审代理律所名称"])             .rename(columns={"执行代理律所名称":"律所名称"})

        dl=pd.concat([dl_1,dl_2,dl_3,dl_4]).drop(columns=["目前有效"])
        return dl    
    
    #不动产，按案件名称整理
    def slct_dyw(self):
        sht_3=pd.read_excel(self.filename,sheet_name=3)

        lst_case=[]
        for i in sht_3["案件名称"]:
            if i not in lst_case:
                lst_case.append(i)

        dyw=[]
        for i in lst_case:
            _dyw=""
            for txt_dyw in sht_3.loc[sht_3["案件名称"]==i]["抵押不动产位置"]:
                _dyw=_dyw+str(txt_dyw)+"；"
            _dyw=_dyw[0:-1]
            dyw.append([i,_dyw])

        lst_dyw = pd.DataFrame(dyw)
        lst_dyw.columns=["案件名称","抵押物情况"]
        lst_dyw=lst_dyw[~lst_dyw["抵押物情况"].isin(["nan"])] #反选，抵押物情况不为nan，即有抵押物情况存在的案件
        
        return lst_dyw
    
    def wt(self):
        print("* 正在写入临时数据... \n")
        wt=pd.ExcelWriter('c:\py\lyy\output.xlsx')
        self.res.to_excel(wt,index=False)
        wt.save()
        print("完成 \n")
    
class wt_excel:
    def __init__(self,filename,outputname="c:\py\lyy\报表"):
        self.filename=filename
        self.outputname=outputname+".xlsx"
    
    def wt(self):
        print("* 正在调整格式... \n")
        wb = opx.load_workbook(self.filename)
        sht = wb['Sheet1']
        sht.insert_rows(1,2)
        mrows=sht.max_row
        mcols=sht.max_column
    
        e="C"+str(mrows)
        
        n=4 #表格数据从第二行开始
        lst=[]
        lst_grp=[]
        for rows in sht["C4":e]:
            for cells in rows:
                lst_grp.append([cells.value,n])
                if cells.value not in lst:
                    lst.append(cells.value)                
                n+=1 
                
        g=[]               
        for i in lst:
            g_0=[] 
            for j in lst_grp:
                if j[0]==i:
                    g_0.append(j[1])
            g.append(g_0)
            
        gp=[]
        col_to_merge=["C","E","F","G","H","K","L"] #需要合并的列坐标
        for i in g:
            if i[0]!=i[-1]:
                for j in col_to_merge:
                    gp.append(j+str(i[0])+":"+j+str(i[-1])) 
          
        for i in gp:
            sht.merge_cells(i)
            
        sht.merge_cells("M2:V2")
        sht.merge_cells("W2:Y2")
        sht.merge_cells("Z2:AC2")
        sht.merge_cells("A1:AC1")
        sht["M2"]="一审"
        sht["W2"].value="二审"
        sht["Z2"].value="执行"
        sht["A1"].value="南宁市南方融资担保有限公司诉讼案件总表"
            
        
            
#         self.adjust_fmt()
            
            
#     def adjust_fmt(self):

        #调整格式     
        title_A_L=sht["A3:L3"]
        
        cols_to_merge=["A","B","C","D","E","F","G","H","I","J","K","L"]
        for i in cols_to_merge:
            sht.merge_cells(i+"2:"+i+"3")
            
        n=0
        for i in sht["A2:L2"][0]:
            i.value=title_A_L[0][n].value
#             print(i.value)
            n+=1
            
        #打包样式
        line_t = Side(style='thin', color='000000')  # 细边框
        line_m = Side(style='medium', color='000000')  # 粗边框
        num_fmt="0.00"
        border0 = Border(top=line_m, bottom=line_m, left=line_m, right=line_m)
        border1 = Border(top=line_t, bottom=line_t, left=line_t, right=line_t)        
        align = Alignment(horizontal='left',vertical='center',wrap_text=True)
        align_title=Alignment(horizontal='center',vertical='center',wrap_text=True)
        sty1 = NamedStyle(name='sty1', border=border1, alignment=align)
        sty_title= NamedStyle(name='sty_title',border=border0,alignment=align_title,font=Font(bold=True,size=11))
        sty_big_title= NamedStyle(name='sty_big_title',alignment=align_title,font=Font(bold=True,size=22))
        sty_num=NamedStyle(name='sty_num',border=border1, alignment=align,number_format=num_fmt)
        
        sht["A1"].style=sty_big_title
        sht["M2"].style=sty_title
        sht["W2"].style=sty_title
        sht["Z2"].style=sty_title
        
        #标题格式
        for r in range(2,4):
            for c in range(1,mcols+1):
                sht.cell(r,c).style=sty_title
        #正文格式
        for r in range(4,mrows+1):
            for c in range(1,mcols+1):
                sht.cell(r,c).style=sty1    
                
        #D、E列格式（数字）
        for r in range(4,mrows+1):
            for c in range(4,6):
                sht.cell(r,c).style=sty_num    
        
        #调整列宽
        #获取每一列的内容的最大宽度
        i = 0
        col_width=[]
        for col in sht.columns:
            for j in range(len(col)):
                if j == 0:
                    col_width.append(len(str(col[j].value)))
                else:
                    # 获得每列中的内容的最大宽度
                    if col_width[i] < len(str(col[j].value)):
                        col_width[i] = len(str(col[j].value))
            i = i + 1

         #设置列宽
        for i in range(len(col_width)):
             # 根据列的数字返回字母
            col_letter = get_column_letter(i+1)
             # 当宽度大于100，宽度设置为100
            if col_width[i] > 100:
                sht.column_dimensions[col_letter].width = 100
             # 只有当宽度大于10，才设置列宽
            elif col_width[i] > 15:
                sht.column_dimensions[col_letter].width = col_width[i] + 2
        
        #调整行高 
        sht.row_dimensions[1].height = 44
        sht.row_dimensions[2].height = 33
        sht.row_dimensions[3].height = 27
        
        print("完成 \n * 正在生成文件：{}... \n".format(self.outputname))
            
        wb.save(self.outputname)
        print("完成\n")
        a=input("按回车退出")
                        
if __name__=="__main__":
    luo=sht_merge("c:\py\lyy\南宁市南方融资担保有限公司诉讼案件总表.xlsm")
    luo.rd()
    luo.wt()
    
    wt=wt_excel("c:\py\lyy\output.xlsx")
    wt.wt()
