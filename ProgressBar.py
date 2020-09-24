import time
def total(i):
    n=100
    k=i/n
    total=30
    jing='#'*round(k*total)
    line='-'*(30-round(k*total))
    out='完成{}%  [{}]'.format(round(k*100),jing+line)    
    print(out,end='\r',flush=True)    
    time.sleep(0.1)    