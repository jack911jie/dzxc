import os
import json

def readConfig(fn):
    with open(fn,'r',encoding='utf-8') as f:
        lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=json.loads(_line)
    return config

def code_to_str(ss):
    s=ss['keywords']
    if isinstance(s,list):
        out=[]
        for i in s:
            out.append(i.decode('utf-8'))
    else:
        out=[ss.decode('utf-8')]
   
#     print(out)
    return out