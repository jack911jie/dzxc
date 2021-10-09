import os
import sys

class TargetDirs:
	def __init__(self,dir='e:\\temp\\照片打标'):
		self.dir=dir
		self.rar=r'C:/"Program Files"/WinRAR/WinRAR.exe'

	def to_rar(self):
		os.chdir(self.dir)
		for fn in os.listdir(self.dir):
			target='"'+fn+'.rar"'
			source='"'+fn+'"'
			zip_command = r'{0} a {1} {2}'.format(self.rar,target,source) 
			os.system(zip_command)


if __name__=='__main__':
	tg=TargetDirs()
	tg.to_rar()