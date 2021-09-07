import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import composing
import readConfig
import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
from PIL import Image,ImageDraw,ImageFont
import numpy as np



class EveryWeek:
	def __init__(self,crs_name='正面管教',tch_name='杨芳芳'):
		self.crs_name=crs_name
		self.tch_name=tch_name
		config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'psychology.config'))
		self.crs_dir=config['课程内容文件夹']
		self.public_dir=config['公共素材']
		self.tch_dir=config['老师素材']
		self.crs_dir=self.crs_dir.replace('$',self.crs_name)
		
		self.book_info=readConfig.readConfig(os.path.join(self.crs_dir,'课程信息',crs_name+'.txt'))
		self.bookcover_src=os.path.join(self.crs_dir,'课程信息',crs_name+'.jpg')

		self.tch_info=readConfig.readConfig(os.path.join(self.tch_dir,'老师信息',tch_name+'.txt'))
		# print(self.tch_info)

	def read_excel(self):
		xls=self.crs_name+'课程信息表.xlsx'
		df=pd.read_excel(os.path.join(self.crs_dir,xls))
		return df

	def exp_pic(self,template_id='K001',class_no=1):
		df=self.read_excel()

		def color_list():
			color={
				'bg':{
					'all':'#FFFFFF',
					'title':'#00807D',
					'small_title':'#effffb',
					'behavior':'#ffffff',
					'book':'#fff7ef',
					'analysis':'#ffffff',
					'exp':'#fff7ef',
					'exp_small':'#ffffff',
					'bottom':'#00807D'
				},
				'edge':{
					'title':'#ffffff',
					'small_title':'#ffffff',
					'behavior':'#999999',
					'book':'#ffffff',
					'analysis':'#00807D',
					'exp':'#ffffff',
					'exp_small':'#d58e00',
					'bottom':'#ffffff'
				},
				'font':{
					'title':'#ffffff',
					'small_title':'#00807D',
					'behavior':'#999999',
					'book':'#656463',
					'analysis':'#002e19',
					'exp':'#d58e00',
					'exp_box_title':'#e6cdb3',
					'tch_intro':'#d58e00',
					'bottom':'#ffffff'
				}
			}

			return color

		def txts():
			title='大智家长课'
			small_title=df[df['模板编号']==template_id]['问题简述'].values[0]
			behavior=df[df['模板编号']==template_id]['问题及场景'].values[0]
			analysis=df[df['模板编号']==template_id]['问题分析'].values[0]
			exp=df[df['模板编号']==template_id]['专家点评'].values[0]
			
			return {'title':title,'small_title':small_title,'behavior':behavior, \
					'analysis':analysis,'exp':exp}

		def ht_cal():
			txt=txts()

			wide=720
			ht_title=220
			ht_small_title=100
			ht_book=300
			ht_bottom=300
			ht_gap=10

			ftsize_small_title=36
			ftsize_behavior=30
			dat_behavior=composing.split_txt_Chn_eng(wid=700,font_size=ftsize_behavior,txt_input=txt['behavior'],Indent='no')
			txt_behavior=dat_behavior[0]
			ht_behavior=int(dat_behavior[1]*ftsize_behavior*2.5)
			
			ftsize_analysis=30
			dat_analysis=composing.split_txt_Chn_eng(wid=700,font_size=ftsize_analysis,txt_input=txt['analysis'],Indent='no')
			txt_analysis=dat_analysis[0]
			ht_analysis=int(dat_analysis[1]*ftsize_analysis*2.3)

			ftsize_exp=30
			dat_exp=composing.split_txt_Chn_eng(wid=700,font_size=ftsize_exp,txt_input=txt['exp'],Indent='no')
			txt_exp=dat_exp[0]
			ht_exp=int(dat_exp[1]*ftsize_exp*2.2)
			ht_exp+=300

			total_ht=ht_title+ht_small_title+ht_behavior+ht_book+ht_analysis+ht_exp+ht_bottom+ht_gap*6

			ht={
				'wide':{
					'total_wt':wide,
					'total':wide,
					'small_title':650,
					'behavior':650,
					'book':650,
					'analysis':650,
					'exp':650,
					'bottom':wide
				},
				'ht':{
					'total_ht':total_ht,
					'bg_title':ht_title,
					'bg_small_title':ht_small_title,
					'bg_behavior':ht_behavior,
					'bg_book':ht_book,
					'bg_analysis':ht_analysis,
					'bg_exp':ht_exp,
					'bg_bottom':ht_bottom,
					'gap':ht_gap
				},
				'ftsize':{
					'small_title':ftsize_small_title,
					'behavior':ftsize_behavior,
					'exp':ftsize_exp,
					'analysis':ftsize_analysis
				},
				'indent_small_box':30
			}

			return {'ht':ht,'txt':txt}

		def draw_bg():
			ht_txt_data=ht_cal()
			size=ht_txt_data['ht']
			# print(size)
			color=color_list()

			#大标题
			p_title=(0,
					0,
					size['wide']['total_wt'],
					size['ht']['bg_title'])

			
			#小标题			
			p_small_title=(p_title[0]+size['indent_small_box'],
							p_title[3]+size['ht']['gap'], 
							p_title[0]+size['indent_small_box']+size['wide']['small_title'],
							p_title[3]+size['ht']['gap']+size['ht']['bg_small_title'])

			#行为描述
			p_behavior=(p_small_title[0], 
						p_small_title[3]+size['ht']['gap'], 
						p_small_title[0]+size['wide']['behavior'], 
						p_small_title[3]+size['ht']['gap']+size['ht']['bg_behavior']
						)
			p_behavior=tuple(p_behavior)

			#书
			p_book=(p_behavior[0], 
				p_behavior[3]+size['ht']['gap']*3, 
				p_behavior[0]+size['wide']['book'], 
				p_behavior[3]+size['ht']['gap']+size['ht']['bg_book']
				)
					
			#分析
			p_analysis=(p_book[0], 
				p_book[3]+size['ht']['gap']*3, 
				p_book[0]+size['wide']['analysis'], 
				p_book[3]+size['ht']['gap']+size['ht']['bg_analysis']
				)
			#专家点评
			p_exp=(p_analysis[0], 
				p_analysis[3]+size['ht']['gap']*3, 
				p_analysis[0]+size['wide']['exp'], 
				p_analysis[3]+size['ht']['gap']+size['ht']['bg_exp']
				)
			p_exp_small=(p_exp[0]+10, 
				p_exp[3]-260, 
				p_exp[2]-10, 
				p_exp[3]-10
				)
			#底部
			p_bottom=(0, 
				p_exp[3]+size['ht']['gap']*3, 
				p_exp[0]+size['wide']['bottom'], 
				p_exp[3]+size['ht']['gap']+size['ht']['bg_bottom']
				)

			bg=Image.new('RGBA',(size['wide']['total_wt'],size['ht']['total_ht']),color=color['bg']['all'])
			draw=ImageDraw.Draw(bg)		
			draw.rectangle(p_title,fill=color['bg']['title'])
			draw.rectangle(p_small_title,fill=color['bg']['small_title'])
			draw.rounded_rectangle(p_behavior,20,fill=color['bg']['behavior'],width=3,outline=color['edge']['behavior'])
			draw.rectangle((p_behavior[0]+15,p_behavior[1]+25,p_behavior[0]+15+10,p_behavior[3]-25),fill=color['edge']['behavior'])
			
			draw.rectangle(p_book,fill=color['bg']['book'])
			draw.rounded_rectangle(p_analysis,20,fill=color['bg']['analysis'],width=3,outline=color['edge']['analysis'])
			draw.rectangle(p_exp,fill=color['bg']['exp'])
			draw.rounded_rectangle(p_exp_small,20,fill=color['bg']['exp_small'],width=3,outline=color['edge']['exp_small'])
			draw.rectangle(p_bottom,fill=color['bg']['bottom'])

			#图片
			#书
			_book_cover=Image.open(self.bookcover_src)
			book_cover=_book_cover.resize((_book_cover.size[0]*250//_book_cover.size[1],250))
			bg.paste(book_cover,(p_book[0]+85,p_book[1]+15))

			#老师
			try:
				_pic_tch=Image.open(os.path.join(self.tch_dir,'老师照片',self.tch_name+'.jpg'))			
			except FileNotFoundError as e:
				_pic_tch=Image.open(os.path.join(self.tch_dir,'老师照片','model.jpg'))

			pic_tch=_pic_tch.resize((_pic_tch.size[0]*150//_pic_tch.size[1],150))
			bg.paste(pic_tch,(p_exp[0]+60,p_exp[3]-150-50))

			#底部二维码
			_logo=Image.open(os.path.join(self.public_dir,'logoForPic.png'))
			logo=_logo.resize((_logo.size[0]*200//_logo.size[1],200))
			a_logo=logo.split()[3]
			bg.paste(logo,(p_bottom[0]+50,p_bottom[1]+50),mask=a_logo)

			_qrcode=Image.open(os.path.join(self.public_dir,'大智小超二维码.jpg'))
			qrcode=_qrcode.resize((_qrcode.size[0]*200//_qrcode.size[1],200))
			bg.paste(qrcode,(p_bottom[0]+450,p_bottom[1]+50))


			#文字部分
			txts=ht_txt_data['txt']
			draw.text((145,35), '大智家长课', fill = color['font']['title'],font=composing.fonts('华康海报体W12(p)',90))  #课程名称 
			draw.text((255,155), '-- 第 '+str(class_no)+' 课 --', fill = color['font']['title'],font=composing.fonts('华康海报体W12(p)',40))  

			if len(txts['small_title'])*size['ftsize']['small_title']>size['wide']['small_title']:
				pass
			else:
				x_small_title=p_small_title[0]+((size['wide']['small_title']-len(txts['small_title'])*size['ftsize']['small_title'])//2)+5
			# print(x_small_title)
			draw.text((x_small_title,p_small_title[1]+20), txts['small_title'], fill = color['font']['small_title'],font=composing.fonts('微软雅黑',size['ftsize']['small_title'])) 
			composing.put_txt_img(draw=draw,
									tt=txts['behavior'],
									total_dis=int(size['wide']['small_title']*0.85),
									xy=[p_behavior[0]+40,p_behavior[1]*1.1],
									dis_line=int(size['ftsize']['behavior']*0.6),
									fill=color['font']['behavior'],
									font_name='楷体',
									font_size=size['ftsize']['behavior'],
									addSPC='add_2spaces')

			x_book_dis=400
			draw.text((p_book[0]+300+(x_book_dis-len(self.book_info['书名'])*60)//2,p_book[1]+45),self.book_info['书名'],fill = color['font']['book'],font=composing.fonts('微软雅黑',50))
			draw.text((p_book[0]+290+(x_book_dis-(len(self.book_info['国家'])+2)*30)//2,p_book[1]+140),'【'+self.book_info['国家']+'】',fill = color['font']['book'],font=composing.fonts('微软雅黑',28))
			draw.text((p_book[0]+340+(x_book_dis-len(self.book_info['作者'])*40)//2,p_book[1]+190),self.book_info['作者'],fill = color['font']['book'],font=composing.fonts('微软雅黑',32))
			
			composing.put_txt_img(draw=draw,
									tt=txts['analysis'],
									total_dis=int(size['wide']['analysis']*0.85),
									xy=[p_analysis[0]+40,p_analysis[1]*1.02],
									dis_line=int(size['ftsize']['analysis']*0.6),
									fill=color['font']['analysis'],
									font_name='楷体',
									font_size=size['ftsize']['analysis'],
									addSPC='add_2spaces')
			
			composing.put_txt_img(draw=draw,
									tt=txts['exp'],
									total_dis=int(size['wide']['exp']*0.86),
									xy=[p_exp[0]+40,p_exp[1]*1.02],
									dis_line=int(size['ftsize']['exp']*0.6),
									fill=color['font']['exp'],
									font_name='楷体',
									font_size=size['ftsize']['exp'],
									addSPC='add_2spaces')
			x_tch_dis=400
			exp_title='大智家长课特约心理专家'
			draw.text(((size['wide']['exp']-len(exp_title)*30)//2+size['indent_small_box'],p_exp[3]-255),exp_title,fill = color['font']['exp_box_title'],font=composing.fonts('微软雅黑',30))
			draw.text((p_exp[0]+220+(x_tch_dis-len(self.tch_info['姓名'])*30)//2,p_exp[3]-195),self.tch_info['姓名'],fill = color['font']['tch_intro'],font=composing.fonts('微软雅黑',30))
			draw.text((p_exp[0]+220+(x_tch_dis-len(self.tch_info['姓名'])*30)//2,p_exp[3]-140),self.tch_info['title'],fill = color['font']['tch_intro'],font=composing.fonts('微软雅黑',20))
			


			# def put_txt_img(draw,tt,total_dis,xy,dis_line,fill,font_name,font_size,addSPC='None')
			bg.show()
		draw_bg()


if __name__=='__main__':
	card=EveryWeek(crs_name='正面管教',tch_name='杨芳芳')
	# card.read_excel(xls='正面管教课程信息表.xlsx')
	card.exp_pic(template_id='K002',class_no=12)