import os,sys,time
import datetime
from datetime import datetime as dt 
from datetime import timedelta as td
print('自動化登入維持程序進行中，請勿關閉視窗')
#sys.argv[1] = 'D:\\python\\autoorder\\login.bat'
#sys.argv[2] = '1200'

os.system(sys.argv[1])
while True:
    nowTime = dt.strftime(dt.now(), '%Y-%m-%d %H:%M:%S')
    nextTime = dt.strftime(dt.now()+td(seconds=int(sys.argv[2])), '%Y-%m-%d %H:%M:%S')
    print('自動化登入維持程序進行中，請勿關閉視窗')
    print('執行時間: '+nowTime+',下次執行: '+nextTime)
    time.sleep(int(sys.argv[2]))
    #os.system('D:\\python\\autoorder\\login.bat')
    os.system(sys.argv[1])
    