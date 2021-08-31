#UI를 생성하고 event에 맞는 함수를 실행한다.
import time
import webbrowser
import tkinter
import core
import dbcontrol
import win32com.client as win32
import os
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import sys

class UI(tkinter.Tk):
    def __init__(self, q):
        self.q = q
        tkinter.Tk.__init__(self)
        self.text = tkinter.Text(self)
        self.text.place(x=100, y=200)
        self.title("hwp active controller")
        self.geometry("740x600+100+100")
        self.resizable(False, False)

        self.label = tkinter.Label(self, text='SRC : ')
        self.label.place(x=2, y=4)
        # 콤보박스에서 선택한 요소를 담을 변수
        self.tempstrsrc = tkinter.StringVar()
        self.cmbsrc = tkinter.ttk.Combobox(self, height=15, width=45, values=dbcontrol.recentsrc, textvariable=self.tempstrsrc)
        self.cmbsrc.bind("<<ComboboxSelected>>", self.SelectSRCFromCombobox)
        self.cmbsrc.place(x=45, y=5)
        if dbcontrol.recentsrc:
            self.cmbsrc.set(dbcontrol.recentsrc[0])
            core.src_dir = dbcontrol.recentsrc[0]

        # Ask Folder button
        self.btn = tkinter.Button(self, text="...", command=lambda: self.asksrc(self, self.cmbsrc))
        self.btn.place(x=400, y=3)

        self.label = tkinter.Label(self, text='DST : ')
        self.label.place(x=2, y=78)
        # Combobox
        self.tempstrdst = tkinter.StringVar()
        self.cmbdst = tkinter.ttk.Combobox(self, height=15, width=45, values=dbcontrol.recentdst, textvariable=self.tempstrdst)
        self.cmbdst.bind("<<ComboboxSelected>>", self.SelectDSTFromCombobox)
        self.cmbdst.place(x=45, y=80)
        if dbcontrol.recentdst:
            self.cmbdst.set(dbcontrol.recentdst[0])
            core.dst_dir = dbcontrol.recentdst[0]

        # Ask Folder Button
        self.btn = tkinter.Button(self, text="...", command=lambda: self.askdst(self, self.cmbdst))
        self.btn.place(x=400, y=78)

        # Help Document Button
        # userinterface.PopUpHelpDocument()로 하면 버튼을 누르지 않아도 실행되어버림.
        self.btn = tkinter.Button(self, text="Help", width=15, height=2, command=self.PopUpHelpDocument)
        self.btn.place(x=620, y=5)

        # 보안 모듈 다운로드 Button
        self.btn = tkinter.Button(self, text='Security Module Download', command=self.downloadsecuritymodule, height=2,
                             width=20)
        self.btn.place(x=460, y=5)

        # 보안 모듈 가이드 문서
        self.btn = tkinter.Button(self, text='SecurityModuleGuide', command=self.OpenSecurityModuleGuideDoc, height=2,
                             width=17)
        self.btn.place(x=605, y=50)

        print(core.src_dir)
        print(core.dst_dir)

        # 실행 버튼
        self.btn = tkinter.Button(self, text='RUN', command=self.AskExecution,
                             height=4, width=10)
        self.btn.place(x=460, y=50)

        self.printlog()

    #Chrome으로 보안 모듈 다운로드 사이트 열기
    def downloadsecuritymodule(self):
        chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
        url = "https://www.hancom.com/board/devdataView.do?board_seq=47&artcl_seq=4085&pageInfo.page=&search_text=#"
        webbrowser.get(chrome_path).open(url)

    #Help Document Pop Up
    def PopUpHelpDocument(self):
        help = "#hwp 파일을 hwpx로 자동변환 하는 프로그램 입니다.\n\n" \
               "#이 프로그램은 한글이 실행되고 여러 문서가 변환되기 때문에 속도가 빠릅니다.\n\n" \
               "#프로그램 사용법\n\n" \
               "1. 원하는 경로를 선택하고 'select' 버튼을 누르세요.\n\n" \
               "2. RUN 버튼을 누르면 Src 폴더에 있는 파일 구조가 Dst 폴더로 복사됩니다.\n\n" \
               "3. Src 폴더에 있는 모든 .hwp 파일이 .hwpx 파일로 변환되어 Dst 폴더에 저장됩니다.\n\n" \
               "기존의 모든 한글 파일을 닫고 실행해주세요\n\n" \
               "\n\n#제약사항\n\n" \
               "한컴 2010 SE 부터 사용가능한 프로그램입니다.\n\n" \
               "여러 버전의 한컴오피스가 설치되어 있다면 가장 최근에 설치된 것이 실행됩니다.\n\n"
        messagebox.showinfo('Help', help)

    #보안모듈 가이드 문서 열기
    def OpenSecurityModuleGuideDoc(self):
        def resource_path(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)

        smguide = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        smguide.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        smguidepath = resource_path("SecurityModuleGuide.hwp")
        smguide.Open(smguidepath)


    #ComboboxSelected이면 선택된 항목을 소스 폴더 변수에 저장
    def SelectSRCFromCombobox(self, event):
        core.src_dir = event.widget.get()
        print(core.src_dir)

    #ComboboxSelected이면 선택된 항목을 목적지 폴더 변수에 저장
    def SelectDSTFromCombobox(self, event):
        core.dst_dir = event.widget.get()
        print(core.dst_dir)

    # SRC 폴더를 선택하는 함수
    def asksrc(self, root, cmbsrc):
        root.dirName = filedialog.askdirectory() # SRC 폴더를 선택한다.
        core.src_dir = root.dirName #SRC 폴더
        cmbsrc.set(core.src_dir)

    # DST 폴더를 선택하는 함수
    def askdst(self, root, cmbdst):
        root.dirName = filedialog.askdirectory()
        core.dst_dir = root.dirName #DST 폴더
        cmbdst.set(core.dst_dir)

    def printlog(self):
        if core.is_done:
            if not self.q.empty():

                # print('flushed')
                while not self.q.empty():
                    temp = self.q.get(0)
                    # print(temp)
                    self.text.insert(tkinter.END, temp + '\n')
                    self.text.see(tkinter.END)
                self.text.insert(tkinter.END, "###########Result###########" + '\n')
                self.text.insert(tkinter.END, "{}/{} 개 변환 완료.".format(core.hwpxcount, core.hwpcount) + '\n')
                self.text.insert(tkinter.END, "열기 실패 : {} 개".format(core.failhwpcount) + '\n')
                self.text.insert(tkinter.END, "변환 실패 : {} 개".format(core.failhwpxcount) + '\n')
                if core.failhwpcountlist:
                    self.text.insert(tkinter.END, "\n---오픈 실패한 파일---\n")
                    for i in range(core.failhwpcount):
                        self.text.insert(tkinter.END, core.failhwpcountlist[i] + '\n')
                if core.failhwpxcountlist:
                    self.text.insert(tkinter.END, "\n---변환 실패한 파일---\n")
                    for i in range(core.failhwpxcount):
                        self.text.insert(tkinter.END, core.failhwpxcountlist[i] + '\n')
                if core.for_distribution:
                    self.text.insert(tkinter.END, "\n---변환 실패한 파일 중 배포용 문서---\n")
                    for i in range(core.for_distribution_count):
                        self.text.insert(tkinter.END, core.for_distribution[i] + '\n')
                if core.timeoutlist:
                    self.text.insert(tkinter.END, "\n---응답시간 초과된 파일---\n")
                    for i in range(core.timeoutcount):
                        self.text.insert(tkinter.END, core.timeoutlist[i] + '\n')
                if core.false_but_converted:
                    self.text.insert(tkinter.END, "\n---삭제 실패한 파일---\n")
                    for i in range(core.false_but_converted_count):
                        self.text.insert(tkinter.END, core.false_but_converted[i] + '\n')
                self.text.insert(tkinter.END, "##########Completed#########" + '\n')
                self.text.see(tkinter.END)

        self.text.after(1000, self.printlog)

    def ClearText(self):
        self.text.delete(0.0, tkinter.END)

    def AskExecution(self):
        askexecution = tkinter.Toplevel(self)
        askexecution.geometry("300x100+400+100")
        label = tkinter.Label(askexecution, text='열려있는 한글 파일이 모두 꺼집니다.\n 정말로 실행하시겠습니까?')
        label.pack()
        btn = tkinter.Button(askexecution, text='Run', width=5, height=2, command=lambda: [askexecution.destroy(), self.ClearText(), core.execution(core.src_dir, core.dst_dir)])
        btn.place(x=125, y=50)


