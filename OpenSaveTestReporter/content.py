#Email에 전송될 content를 가공한다.
import pandas as pd
import sqlite3
import log
import parsing

loggingman = log.loggingman()
conn = sqlite3.connect("RecentData.db")
cur = conn.cursor()

def Create_Table(list, tablename, processornot):
    temp = '(extname TEXT, '
    if processornot:
        for value in list:
            temp += Process_Statusname_ForDB(value[0]) + ' INTEGER, '
    else:
        for value in list:
            temp += value + ' INTEGER, '

    temp = temp[:len(temp)-2] + ')'
    create_table_sql = "CREATE TABLE IF NOT EXISTS {}".format(tablename) + temp
    # print(create_table_sql)

    return create_table_sql

def Process_Statusname_ForDB(str):
    str = str.replace(' ', '_')
    str = str.replace('-', '_')
    str = str.replace('(', '')
    str = str.replace(')', '')

    return str

def ProduceContents(data):
    global count
    count_per_extname = data.pop()
    buildfilepath = data.pop()
    # extname을 추출
    extnames = data.pop()
    # 메일 제목
    titleinfo = data.pop()
    title = '[CI 결과 공유]'
    title += titleinfo[0]
    # extname에 따른 DB에서 테이블의 이름을 담을 변수
    tablename = ""

    # 테이블 생성
    if '.odt' in extnames:
        title += '[{}]'.format('ODT')
        tablename = "ODT"
        for i in range(len(extnames)):
            # status(현상)의 이름을 extname에 따라 바꾼다.
            for key in data[0][i].keys():
                if data[0][i][key][0] == 'MS오픈실패':
                    data[0][i][key][0] = "원본_LO오픈실패"
                elif data[0][i][key][0] == 'MS복구팝업':
                    data[0][i][key][0] = "원본_LO복구팝업"
                elif data[0][i][key][0] == 'MS제한된보기':
                    data[0][i][key][0] = "원본_LO제한된보기"
                else:
                    pass

        cur.execute(Create_Table(parsing.status.values(), tablename, 1))

    elif '.hwp' in extnames:
        title += '[{}]'.format('HWP')
        tablename = "HWP"
        for i in range(len(extnames)):
            # status(현상)의 이름을 extname에 따라 바꾼다.
            for key in data[0][i].keys():
                if data[0][i][key][0] == 'MS오픈실패':
                    data[0][i][key][0] = "HWP오픈실패"
                elif data[0][i][key][0] == 'MS복구팝업':
                    data[0][i][key][0] = "HWP복구팝업"
                elif data[0][i][key][0] == 'MS제한된보기':
                    data[0][i][key][0] = "HWP제한된보기"
                else:
                    pass

        cur.execute(Create_Table(parsing.status.values(), tablename, 1))

    elif '.ppt' in extnames or '.pptx' in extnames:
        title += '[{}]'.format('Slide')
        tablename = "SLIDE"

        cur.execute(Create_Table(parsing.status.values(), tablename, 1))

    elif '.doc' in extnames or '.docx' in extnames:
        doc_count = 0
        docx_count = 0

        if '.doc' in count_per_extname.keys():
            doc_count = count_per_extname['.doc']
        else:
            pass

        if '.docx' in count_per_extname.keys():
            docx_count = count_per_extname['.docx']
        else:
            pass

        if doc_count <= docx_count:
            title += '[{}]'.format('DOCX')
            tablename = 'DOCX'

            cur.execute(Create_Table(parsing.status.values(), tablename, 1))
        else:
            title += '[{}]'.format('DOC')
            tablename = 'DOC'

            cur.execute(Create_Table(parsing.status.values(), tablename, 1))
    elif '.xls' in extnames or '.xlsx' in extnames:
        title += '[{}]'.format('Sheet')
        tablename = "SHEET"

        cur.execute(Create_Table(parsing.status.values(), tablename, 1))
    elif '.pdf' in extnames:
        title += '[{}]'.format('PDF')
        tablename = "PDF"

        cur.execute(Create_Table(parsing.status.values(), tablename, 1))
    else:
        pass

    title += titleinfo[1]
    title += titleinfo[2]

    # 표의 Column 이름
    col = ['현상']

    # 현상의 이름을 추출
    statusname = []
    for value in data[0][0].values():
        statusname.append(value[0])

    temp = []

    # 이전 데이터를 받을 리스트.
    previous_data = []

    cur.execute("SELECT * FROM " + tablename)
    dbcolumns = [desc[0] for desc in cur.description]
    row = cur.fetchall()
    if row:
        previous_data.append(statusname)  # 표의 형식에 맞추기 위해 현상이름을 추가한다.
        for i in range(len(row)):
            item = list(row[i])
            extname = item.pop(0)
            # extname 인 것만 col에 넣어준다.
            if extname != 'SUM':
                col.append(extname)

            temp_to_append_previous_data = []
            # 데이터베이스의 이전 데이터의 합계
            total_count_prev = 0
            # 데이터베이스의 이전 데이터의 실제 이슈 합계
            real_count_prev = 0

            # parsing.status를 dbcolumns와 비교하기 위해 같은 형식으로 맞추기 위한 리스트
            temp_parsing_status_name = [] # 현상의 이름만 담는 리스트
            temp_parsing_status_realornot = [] # 실제 이슈 합계에 포함되는지 여부만 담는 리스트

            # dbcolumns는 확장자의 이름에 따라 바뀌지 현상이름이 바뀌지 않음(예를 들면, odt는 원본_LO복구팝업인데 MS복구팝업으로 표시됨).
            # 따라서 원본 현상이름을 가지고 와서 비교한다.
            # 하지만 표에 표기할 때에는 바뀐 이름으로 들어가므로 상관없음.
            for name in parsing.status.values():
                temp_parsing_status_name.append(Process_Statusname_ForDB(name[0]))
            for value in parsing.status.values():
                temp_parsing_status_realornot.append(value[1])

            # db의 컬럼에 status에 해당되는 것이 있으면 리스트에 추가하고
            # 합계를 증가시키고, 실제 이슈 합계에 해당하면 증가시킨다.
            # dbcolumns의 맨 앞에는 extname이 들어가 있으므로 index는 1부터 시작
            for j in range(len(temp_parsing_status_name)):
                if temp_parsing_status_name[j] in dbcolumns[1:len(dbcolumns)]:
                    ind = dbcolumns.index(temp_parsing_status_name[j])
                    temp_to_append_previous_data.append(row[i][ind])
                    total_count_prev += row[i][ind]
                    if temp_parsing_status_realornot[j]:
                        real_count_prev += row[i][ind]
                else:
                    temp_to_append_previous_data.append(0)

            temp_to_append_previous_data.extend([total_count_prev, real_count_prev])
            previous_data.append(temp_to_append_previous_data)

        if dbcolumns[1:len(dbcolumns)] != temp_parsing_status_name:
            cur.execute('drop table if exists {}'.format(tablename))
            cur.execute(Create_Table(temp_parsing_status_name, tablename , 0))

        if len(row) > 1:
            col.append('합계')
    statusname.extend(['합계', '실제이슈합계'])
    # 데이터를 삭제한다.
    cur.execute("DELETE FROM " + tablename)
    # extname의 순서에 맞게 temp에 list로 저장한다.
    for i in range(len(data[0])):
        col.append(extnames[i])
        count = []
        for value in data[0][i].values():
            count.append(value[2])
        temp.append(count)

        # sql insert 문에 들어갈 와일드카드
        wildcards = '('
        # count와 extname을 삽입하기 위해서
        for j in range(len(count)+1):
            if j == len(count)+1-1:
                wildcards += '?)'
            else:
                wildcards += '?,'
        sql = 'INSERT INTO ' + tablename + ' VALUES' + wildcards
        # print(sql)
        # 삽입될 데이터
        values_to_be_inserted = []
        values_to_be_inserted.append(extnames[i])
        for item in count:
            values_to_be_inserted.append(item)
        cur.execute(sql, tuple(values_to_be_inserted))

    if len(extnames) > 1:
        # Error 별로 집계된 수
        col.append("합계")
        total_per_error = []
        for i in range(len(temp[0])):
            total = 0
            for j in range(0, len(extnames)):
                total += temp[j][i]
                # print(total)
            total_per_error.append(total)
        # sql insert 문에 들어갈 와일드카드
        wildcards = '('
        # total_per_error와 extname(SUM), 합계, 실제이슈합계를 삽입하기 위해서
        for k in range(len(total_per_error) + 1):
            if k == len(total_per_error) + 1 - 1:
                wildcards += '?)'
            else:
                wildcards += '?,'
        sql = 'INSERT INTO ' + tablename + ' VALUES' + wildcards
        values_to_be_inserted = []
        values_to_be_inserted.append('SUM')
        for item in total_per_error:
            values_to_be_inserted.append(item)
        # 합계와 실제이슈합계
        # values_to_be_inserted.extend([0, 0])
        cur.execute(sql, tuple(values_to_be_inserted))
        temp.append(total_per_error)

    # extname별 합계 및 실제 이슈 합계 계산
    real_count = [0 for i in range(len(extnames))]
    total_count = [0 for i in range(len(extnames))]
    for i in range(len(extnames)):
        for value in data[0][i].values():
            total_count[i] += value[2]
            if value[1] == 1:
                real_count[i] += value[2]

    real_count_list_to_append = []
    total_count_list_to_append = []
    # real_count_list_to_append.append('실제 이슈 합계')
    # total_count_list_to_append.append('합계')
    real_count_list_to_append.extend(real_count)
    total_count_list_to_append.extend(total_count)
    if len(extnames) > 1:
        real_count_tot = sum(real_count)
        real_count_list_to_append.append(real_count_tot)

        if real_count_tot == 0:
            title += ' 특이사항 없음'
        else:
            title += ' 이슈 {}개 발생'.format(real_count_tot)
        total_count_list_to_append.append(sum(total_count))

    else:
        if real_count[0] == 0:
            title += ' 특이사항 없음'
        else:
            title += ' 이슈 {}개 발생'.format(real_count[0])

    conn.commit()
    # temp를 transpose해 표의 형태에 맞게 바꾼다.
    table = list(map(list, zip(*temp)))
    table.append(total_count_list_to_append)
    table.append(real_count_list_to_append)

    if previous_data:
        statusname_and_previous_data = list(map(list, zip(*previous_data)))
        for i in range(len(statusname_and_previous_data)):
            statusname_and_previous_data[i].extend(table[i])
    else:
        statusname_and_data = []
        for item in statusname:
            statusname_and_data.append([item])
        for i in range(len(statusname)):
            statusname_and_data[i].extend(table[i])

    # table을 html로 변환한다.
    # 이전 데이터가 존재할 경우.
    if row:
        # '이전', '오늘'을 표시해주는 컬럼.
        previous_today = []
        if len(row) > 1:
            previous_today = ['','이전','이전','이전','오늘','오늘','오늘']
        else:
            previous_today = ['','이전','오늘']
        # print(statusname_and_previous_data)
        # print(previous_today)
        # print(col)
        contents = pd.DataFrame(statusname_and_previous_data, columns=[previous_today,col]).to_html(index=False)
    # 이전 데이터가 존재하지 않을 경우.
    else:
        contents = pd.DataFrame(statusname_and_data, columns=col).to_html(index=False)
    greeting = '안녕하십니까 OpenSaveTestReporter입니다.<br>'
    temp = titleinfo[1][1:len(titleinfo[1])-1]
    greeting += '{} 엔진 : {}<br>'.format(temp, buildfilepath)
    greeting += contents

    return greeting, title
