import userinterface
import core

if __name__ == "__main__":
    UI = userinterface.UI(core.q)
    UI.mainloop()

# # use Tkinter to show a digital clock
# # tested with Python24    vegaseat    10sep2006
# from tkinter import *
# import time
# root = Tk()
# time1 = ''
# clock = Label(root, font=('times', 20, 'bold'), bg='green')
# clock.pack(fill=BOTH, expand=1)
# def tick():
#     global time1
#     # get the current local time from the PC
#     time2 = time.strftime('%H:%M:%S')
#     # if time string has changed, update it
#     if time2 != time1:
#         time1 = time2
#         clock.config(text=time2)
#     # calls itself every 200 milliseconds
#     # to update the time display as needed
#     # could use >200 ms, but display gets jerky
#     clock.after(200, tick)
# tick()
# root.mainloop(  )

# import os
#
# testlist = []
# for path, dirs, files in os.walk("C:\\Users\\dale.h.lee\\Desktop\\1차"):
#     for file in files:
#         testlist.append(file)
#
# copylist = []
# for path, dirs, files in os.walk("C:\\Users\\dale.h.lee\\Desktop\\copy"):
#     for file in files:
#         copylist.append(file)
#
# for i in range(len(testlist)):
#     testlist[i] = testlist[i].replace('.hwp', '')
#     testlist[i] = testlist[i].replace('.HWP', '')
# for i in range(len(copylist)):
#     copylist[i] = copylist[i].replace('.hwpx', '')
#
# print(len(testlist))
# print(len(copylist))
#
# for i in range(319):
#     if testlist[i] not in copylist:
#         print(testlist[i])
# import win32com.client as win32
# import pygetwindow as gw
# import win32gui
# import win32con
# hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
# print(hwp.Open("C:\\Users\\dale.h.lee\\Desktop\\test\\.hwp"))
# # hwnd = gw.getWindowsWithTitle('스크립트 매크로 실행')
# # hwnd[0].close()
# # print(hwnd)
#
# import threading, time
#
# def count(e):
#     for i in range(1000000000):
#         time.sleep(1)
#         a = i
#         if not e.isSet():
#             print(a)
#         else:
#             return
#
#
# if __name__ == '__main__':
#     e = threading.Event()
#     t = threading.Thread(target=count, args=(e,))
#     t.start()
#     t.join(timeout=5)
#     if t.is_alive():
#         print('Thread Still Alive')
#         e.set()
#     else:
#         pass
#
# import time
# import win32gui, win32con
# import pygetwindow as gw
# import win32clipboard
# import win32
# hwnd = gw.getWindowsWithTitle('10118656')[0].title
# if '[배포용 문서]' in hwnd:
#     print(123)
#
# # app = 'Chrome'
# # save = []
# # while True:
# #     time.sleep(2)
# #
# #     hwnd = win32gui.GetForegroundWindow()
# #     print(hwnd)
# #     text = win32gui.GetWindowText(hwnd)
# #     print(text)








