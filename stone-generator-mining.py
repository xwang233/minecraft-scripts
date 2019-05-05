import win32com.client 
import sys 
import time 
import random 
import threading

def main(): 
    print('python version:', sys.version)

    ts = win32com.client.Dispatch('ts.tssoft')

    need_ver = '4.019'
    if ts.ver() != need_ver: 
        print('register failed')
        return 
    print('register successful')

    hwnd = ts.FindWindow('', 'Minecraft 1.14')
    ts_ret = ts.BindWindow(hwnd, 'normal', 'windows', 'normal', 0)
    if ts_ret == 0: 
        print('binding failed')
        return 
    print('binding successful')

    try: 
        ts.LeftDown()
        current = 1
        while current < 10: 
            time.sleep(480)
            ts.WheelDown()
            print('mouse wheel rolled down')
    except: 
        ts.LeftUp() 

    ts.UnBindWindow()
    print('unbinding finished')


if __name__ == "__main__":
    main() 