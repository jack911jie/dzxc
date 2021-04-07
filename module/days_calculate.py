import datetime

def calculate_age(birth_s='20181215'):
    birth_d = datetime.datetime.strptime(birth_s, "%Y%m%d")
    today_d = datetime.datetime.now()
    birth_t = birth_d.replace(year=today_d.year)
    if today_d > birth_t:
        age = today_d.year - birth_d.year
    else:
        age = today_d.year - birth_d.year - 1
    return age


def calculate_days(date_input='20181215'):
    today=datetime.datetime.now()
    date_s = datetime.datetime.strptime(date_input, "%Y%m%d")    
    return (today-date_s).days



def num_to_ch(num):
    if isinstance(num,str):
        wd={'1':'一','2':'二','3':'三','4':'四','5':'五','6':'六','7':'日'}
    elif isinstance(num,int):
        wd={1:'一',2:'二',3:'三',4:'四',5:'五',6:'六',7:'日'}
    else:
        wd='不是星期数'
    return wd[num]