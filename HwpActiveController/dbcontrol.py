#recent.db와 연동하고 최근 접근한 소스 폴더와 목적지 폴더를 db로부터 읽어온다.
#최근 선택한 폴더를 db에 저장하는 메소드를 정의한다.

import os
import sqlite3

# db에서 검색한 최근 접근한 소스 폴더
recentsrc = []
# db에서 검색한 최근 접근한 DST 폴더
recentdst = []

# db connection
conn = sqlite3.connect("recent.db", isolation_level=None)
c = conn.cursor()
# 최근 접근 소스 폴더에 대한 db생성
c.execute("CREATE TABLE IF NOT EXISTS RecentFolderSRC(Path text, AccessTime double)")
# 최근 접근 DST 폴더에 대한 db생성
c.execute("CREATE TABLE IF NOT EXISTS RecentFolderDST(Path text, AccessTime double)")

# recent.db를 읽어 최근 사용 폴더 경로를 가져온다.
# n == 1이면 최근 접근한 소스 폴더 데이터를 가져오고
# n == 2이면 최근 접근한 DST 폴더 데이터를 가져온다.
def readrecentpath(n):
    if n==1:
        c.execute("select Path from RecentFolderSRC order by AccessTime desc")
        recentsrc = []
        temp = c.fetchall()

        for row in temp:
            recentsrc.append(row[0])
        return recentsrc
    elif n==2:
        c.execute("select Path from RecentFolderDST order by AccessTime desc")
        recentdst = []
        temp = c.fetchall()

        for row in temp:
            recentdst.append(row[0])
        return recentdst

# recent.db에 사용한 폴더 경로를 저장한다.
def writerecentpath(SRCTEXT, DSTTEXT):
    c.execute("create trigger if not exists FolderControlSRC "
              "after insert on RecentFolderSRC "
              "when (select count(*) from RecentFolderSRC) > 4 "
              "begin "
              "delete from RecentFolderSRC "
              "where AccessTime in (select min(Accesstime) from RecentFolderSRC); "
              "end;")
    c.execute("create trigger if not exists DuplicateControlSRC "
              "before insert on RecentFolderSRC "
              "begin "
              "delete from RecentFolderSRC "
              "where Path = New.Path; "
              "end;")
    # FolderControlSRC는 4개가 넘으면 가장 오래 전에 접근한 폴더 데이터를 삭제하는 trigger
    # DuplicateControlSRC는 중복된 것이 있으면 새로운 데이터는 삽입하고 존재하던 데이터를 삭제하는 trigger
    AccessTime = os.path.getatime(SRCTEXT) # SRC폴더 최근 접근시간을 가져온다.
    params = (SRCTEXT, AccessTime)
    c.execute("insert into RecentFolderSRC values(?, ?)", params) # db에 SRC폴더 경로와 접근 시간을 삽입

    c.execute("create trigger if not exists FolderControlDST "
              "after insert on RecentFolderDST "
              "when (select count(*) from RecentFolderDST) > 4 "
              "begin "
              "delete from RecentFolderDST "
              "where AccessTime in (select min(Accesstime) from RecentFolderDST); "
              "end;")
    c.execute("create trigger if not exists DuplicateControlDST "
              "before insert on RecentFolderDST "
              "begin "
              "delete from RecentFolderDST "
              "where Path = New.Path; "
              "end;")
    # FolderControlDST는 4개가 넘으면 가장 오래 전에 접근한 폴더 데이터를 삭제하는 trigger
    # DuplicateControlDST는 중복된 것이 있으면 새로운 데이터는 삽입하고 존재하던 데이터를 삭제하는 trigger
    AccessTime = os.path.getatime(DSTTEXT) # DST폴더 최근 접근시간을 가져온다.
    params = (DSTTEXT, AccessTime)
    c.execute("insert into RecentFolderDST values(?, ?)", params) # db에 DST폴더 경로와 접근 시간을 삽입

def ClearDB():
    c.execute("delete from recentfoldersrc where accesstime > 0")
    c.execute("delete from recentfolderdst where accesstime > 0")

#ClearDB()
recentsrc = readrecentpath(1)
recentdst = readrecentpath(2)
