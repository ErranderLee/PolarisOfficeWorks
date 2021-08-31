#파일 변환 함수를 포함한다.
#디렉토리를 탐색하면서 파일 변환을 수행하는 함수를 포함한다.
#UI에서 RUN 버튼을 누를 때 실행되는 함수를 포함한다.

import os
import time
import pythoncom
import dbcontrol
import win32com.client as win32
import threading
import queue
import pygetwindow as gw
import win32gui
import win32con

# SRC 폴더
src_dir = ""

# DST 폴더
dst_dir = ""

#변환될 파일의 수
hwpcount = 0

#변환된 파일의 수
hwpxcount = 0

#열기 실패한 파일의 수와 리스트
failhwpcount = 0
failhwpcountlist = []

#변환 실패한 파일의 수와 리스트
failhwpxcount = 0
failhwpxcountlist = []

#시간 초과로 인한 오픈 실패
timeoutcount = 0
timeoutlist = []

# SaveAs의 return 값이 false 지만 저장이 된 파일
false_but_converted_count = 0
false_but_converted = []

# 배포용 문서
for_distribution_count = 0
for_distribution = []

#GUI에 출력될 log
log = ""

#log를 담을 Queue
q = queue.Queue()

#팝업창 이름
#우선순위에 따라서 리스트에 넣어야 함.
#예를 들어 '한글'이 '스크립트' 보다 앞에 있으면 RPC 서버 통신 오류 발생
popup_name = ['스크립트', '한글']

# 모든 파일에 대한 작업이 끝났음을 알리는 변수
is_done = 0

def HwpOpen(src_dir, dst_dir, filename, e):
    global hwpxcount, failhwpcount, failhwpxcount, false_but_converted_count, for_distribution_count,\
        failhwpcountlist, failhwpxcountlist, for_distribution, false_but_converted, log

    try:
        #hwp connection
        pythoncom.CoInitialize()
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        #file open
        openresult = hwp.Open(os.path.join(src_dir, filename))

        if openresult:
            log = "{}를 한컴으로 열었습니다.".format(filename)
            q.put(log)

            if not e.isSet():
                if hwp.SaveAs(dst_dir + "\\" + filename[:len(filename) - 4] + ".hwpx", "HWPX"):
                    log = "{}를 변환 완료했습니다.".format(filename)
                    q.put(log)
                    hwpxcount = hwpxcount + 1
                else:  # 변환 실패
                    log = '{} 변환을 실패했습니다.'.format(filename)
                    q.put(log)
                    failhwpxcount = failhwpxcount + 1
                    failhwpxcountlist.append(os.path.join(src_dir, filename))

                    if hwp.EditMode == 17:
                        for_distribution_count += 1
                        for_distribution.append(os.path.join(src_dir, filename))

                    failed_but_converted = os.path.join(dst_dir, filename)+'x'
                    if os.path.isfile(failed_but_converted):
                        os.remove(failed_but_converted)
                        log = '{} 손상되어 변환된 파일을 삭제합니다.'.format(failed_but_converted)
                        q.put(log)
                        if not os.path.isfile(failed_but_converted):
                            log = '{} 손상되어 변환된 파일을 삭제했습니다.'.format(failed_but_converted)
                            q.put(log)
                        else:
                            false_but_converted_count += 1
                            false_but_converted.append(failed_but_converted)
                            log = '{} 손상되어 변환된 파일의 삭제를 실패했습니다.'.format(failed_but_converted)
                            q.put(log)
            else:
                # print('SET')
                log = "{} 열기 실패 했습니다.".format(filename)
                q.put(log)
                failhwpcount = failhwpcount + 1
                failhwpcountlist.append(os.path.join(src_dir, filename))
        else:  # 열기 실패
            log = "{} 열기 실패 했습니다.".format(filename)
            q.put(log)
            failhwpcount = failhwpcount + 1
            failhwpcountlist.append(os.path.join(src_dir, filename))
        log = "{}/{} 개 변환 완료.".format(hwpxcount, hwpcount)
        q.put(log)
    except:
        # 팝업창이 뜨지 않는 문서인데 일정 시간 기다려도 켜지지 않으면 가장 앞에 떠 있는
        # '한글' 창을 찾아서 종료시키므로 RPC 통신 오류가 발생한다.
        # 따라서 except 문으로 넘어오게 되는데, event 객체의 flag가 set 되어 있으면 시간초과로 하고 넘어간다.
        if e.isSet():
            timeoutcount += 1
            timeoutlist.append(os.path.join(src_dir, filename))
            log = "{} 시간 초과로 열기 실패 했습니다.".format(filename)
            q.put(log)
            log = "{}/{} 개 변환 완료.".format(hwpxcount, hwpcount)
            q.put(log)
        else:
            log = "{} 에서 오류 발생으로 다시 시도합니다.".format(filename)
            q.put(log)
            time.sleep(2)
            pythoncom.CoInitialize()
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            # file open
            openresult = hwp.Open(os.path.join(src_dir, filename))

            if openresult:
                log = "{}를 한컴으로 열었습니다.".format(filename)
                q.put(log)
                if not e.isSet():
                    if hwp.SaveAs(dst_dir + "\\" + filename[:len(filename) - 4] + ".hwpx", "HWPX"):
                        log = "{}를 변환 완료했습니다.".format(filename)
                        q.put(log)
                        hwpxcount = hwpxcount + 1
                    else:  # 변환 실패
                        log = '{} 변환을 실패했습니다.'.format(filename)
                        q.put(log)
                        failhwpxcount = failhwpxcount + 1
                        failhwpxcountlist.append(os.path.join(src_dir, filename))

                        if hwp.EditMode == 17:
                            for_distribution_count += 1
                            for_distribution.append(os.path.join(src_dir, filename))

                        failed_but_converted = os.path.join(dst_dir, filename) + 'x'
                        if os.path.isfile(failed_but_converted):
                            os.remove(failed_but_converted)
                            log = '{} 손상되어 변환된 파일을 삭제합니다.'.format(failed_but_converted)
                            q.put(log)
                            if not os.path.isfile(failed_but_converted):
                                log = '{} 손상되어 변환된 파일을 삭제했습니다.'.format(failed_but_converted)
                                q.put(log)
                            else:
                                false_but_converted_count += 1
                                false_but_converted.append(failed_but_converted)
                                log = '{} 손상되어 변환된 파일의 삭제를 실패했습니다.'.format(failed_but_converted)
                                q.put(log)
                else:
                    # print('SET')
                    log = "{} 열기 실패 했습니다.".format(filename)
                    q.put(log)
                    failhwpcount = failhwpcount + 1
                    failhwpcountlist.append(os.path.join(src_dir, filename))
            else:  # 열기 실패
                log = "{} 열기 실패 했습니다.".format(filename)
                q.put(log)
                failhwpcount = failhwpcount + 1
                failhwpcountlist.append(os.path.join(src_dir, filename))
            log = "{}/{} 개 변환 완료.".format(hwpxcount, hwpcount)
            q.put(log)

    finally:
        try:
            hwp.Quit()
            pythoncom.CoUninitialize()
        except:
            log = "{} 닫기 오류 발생하여 재시도 합니다.".format(filename)
            q.put(log)
            time.sleep(2)
            hwp.Quit()
            pythoncom.CoUninitialize()



#SRC 폴더에서 hwp 확장자를 찾아서 hwpx로 변환하여 DST 폴더에 저장하는 함수
def hwptohwpx(src_dir, dst_dir):
    global hwpxcount, failhwpcount, failhwpxcount, failhwpcountlist, failhwpxcountlist
    for i in os.listdir():
        temp = i[len(i) - 4:].lower()
        if temp == '.hwp':
            e = threading.Event()
            # HwpOpen을 Thread로 실행
            thread = threading.Thread(target=HwpOpen, args=(src_dir, dst_dir, i, e))
            thread.start()
            # HwpOpen Thread를 최대 30초간 기다림.
            thread.join(timeout=30)
            if thread.is_alive():
                # 닫히지 않은 파일을 닫기 위해 팝업창의 이름이 타이틀에 들어간 윈도우를 가져온다.
                for name in popup_name:
                    hwnd = gw.getWindowsWithTitle(name)
                    if hwnd:
                        hwnd = hwnd[0]
                        hwnd.close()
                        break
                # event 객체의 플래그를 set
                e.set()
#폴더 내의 폴더를 탐색
def changedir(src_dir, dst_dir):
    # SRC 폴더로 이동
    os.chdir(src_dir)
    #hwp to hwpxw
    hwptohwpx(src_dir, dst_dir)
    #find child dir
    directories = []
    #하위 폴더의 이름을 저장한다.
    for fold in os.listdir(src_dir):
        if os.path.isdir(os.path.join(src_dir, fold)):
            directories.append(fold)
    #rec call
    for i in range(len(directories)):
        changedir(src_dir + '\\' + directories[i], dst_dir + '\\' + directories[i])

#hwp 파일 수 파악
def CountHwp(src_dir):
    global hwpcount
    for path, dirs, files in os.walk(src_dir):
        for file in files:
            if os.path.splitext(file)[1].lower() == '.hwp':
                hwpcount = hwpcount + 1

#실행하는 함수
def execution(src_dir, dst_dir):
    # 기능을 실행하기 전 열려있는 모든 한글 파일을 닫는다.
    os.system('taskkill /f /im Hwp.exe')
    #변환될, 변환된 파일의 수, 열기 실패, 변환 실패한 파일의 수
    global hwpcount, hwpxcount, failhwpcount, failhwpxcount, failhwpcountlist, failhwpxcountlist, \
        hwp, timeoutcount, timeoutlist, is_done, false_but_converted, false_but_converted_count, \
        for_distribution, for_distribution_count
    hwpcount = 0
    hwpxcount = 0
    failhwpcount = 0
    failhwpxcount = 0
    timeoutcount = 0
    false_but_converted_count = 0
    for_distribution_count = 0
    is_done = 0
    failhwpcountlist = []
    failhwpxcountlist = []
    timeoutlist = []
    false_but_converted = []
    for_distribution = []

    #python에서 읽을 수 있는 형식으로 가공
    src_dir = src_dir.replace("/", "\\")
    dst_dir = dst_dir.replace("/", "\\")
    #robocopy를 통해 폴더 구조만 복사
    os.system('robocopy' + ' "' + src_dir + '" "' + dst_dir + '" ' + '/e /xf *.* /xd' + ' ' + dst_dir)
    print('robocopy' + ' "' + src_dir + '" "' + dst_dir + '" ' + '/e /xf *.* /xd' + ' ' + dst_dir)
    dbcontrol.writerecentpath(src_dir, dst_dir)
    CountHwp(src_dir)
    changedir(src_dir, dst_dir)
    while True:
        if threading.activeCount() == 1:
            is_done = 1
            break