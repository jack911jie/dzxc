import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'module'))
import composing
import readConfig
import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)



class EveryWeek:
	def __init__(self):
		config=readConfig.readConfig(os.path.join(os.path.dirname(__file__),'psychology.config'))
		self.crs_zmgj_dir=config['正面管教内容文件夹']


	def read_excel(self,xls='正面管教课程信息表.xlsx'):
		df=pd.read_excel(os.path.join(self.crs_zmgj_dir,xls))

if __name__=='__main__':
	card=EveryWeek()
	card.read_excel(xls='正面管教课程信息表.xlsx')