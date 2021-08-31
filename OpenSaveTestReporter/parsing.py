#텍스트 파일의 내용을 parsing해 Data를 추출한다.
import copy
import datetime
import os
import log

loggingman = log.loggingman()

# key : value([에러이름, 매칭될 항목 이름]), 표기되어 있지 않은 것은 모두 기타.
errortypes = {1 : ['eErrorCheck_PCOfficeCloseFail', 1], 2 : ['eErrorCheck_LoadFail', 2],
                  3 : ['eErrorCheck_SaveFail', 3], 4 : ['eErrorCheck_ImageDifferent', 50], 5 : ['eErrorCheck_MessageboxCheck', 50],
                  9 : ['eErrorCheck_UnSupportedMacro', 50], 10 : ['eErrorCheck_ChangeReadOnlyModeWithErrorData', 4],
                  11 : ['eErrorCheck_PageMoveFail', 50], 12 : ['eErrorCheck_PCOfficeCrash', 5], 13 : ['eErrorCheck_PostMessageFail', 50],
              98 : ['eErrorCheck_EmptyActionLog', 16], 99 : ['eErrorCheck_Unknown', 17],
              1003 : ['eEngineException_TestTimeOutAndKill', 50], 1004 : ['eEngineException_UnhandledException', 6], 1005 : ['eEngineException_UnhandledTimeoutException', 50],
              1006 : ['eEngineException_Terminated', 7],
              1100 : ['eMSOfficeValidationAbort_TestTimeOutAndKill', 8], 1101 : ['eMSOfficeValidationAbort_PopUp_UnknownBlocking', 50],
              1102 : ['eMSOfficeValidationAbort_PopUp_Thisworkbookcontainslinkstootherdatasources', 9], 1103 : ['eMSOfficeValidationAbort_PopUp_Thereisnotenoughmemoryordiskspacetoupdatethedisplay', 50],
              1104 : ['eMSOfficeValidationAbort_PopUp_RestorePopup', 50], 1105 : ['eMSOfficeValidationAbort_PopUp_DataLossPopUp', 9],
              1106 : ['eMSOfficeValidationAbort_PopUp_FormatLossPopUp', 9], 1107 : ['eMSOfficeValidationAbort_PopUp_StyleExcessPopUp', 9],
              1108 : ['eMSOfficeValidationAbort_PopUp_FormulaLabelPopup', 9], 1109 : ['eMSOfficeValidationAbort_PopUp_CircularReference', 9],
              1110 : ['eMSOfficeValidationAbort_PopUp_WrongName', 9],
              1300 : ['eMSOfficeValidationOpenFail_UnknownReason', 50], 1301 : ['eMSOfficeValidationOpenFail_Doyouwanttorecoverthecontentsofthisworkbook', 10],
              1302 : ['eMSOfficeValidationOpenFail_PowerPointcanattempttorepairthepresentation', 10], 1303 : ['eMSOfficeValidationOpenFail_Wecannotopendocxbecausewefoundaproblemwithitscontents', 9],
              1304 : ['eMSOfficeValidationOpenFail_RestrictedView', 11],
              1400 : ['eConnectionFail_SinceStart', 15], 1401 : ['eConnectionFail_SinceInit', 15],
              1402 : ['eConnectionFail_SinceOpen', 12], 1403 : ['eConnectionFail_SinceBeginMacro', 15],
              1404 : ['eConnectionFail_SinceEndMacro', 15], 1405 : ['eConnectionFail_SinceBeginThumbnail', 15],
              1406 : ['eConnectionFail_SinceEndThumbnail', 15], 1407 : ['eConnectionFail_SinceBeginImagr', 15], 1408 : ['eConnectionFail_SinceEndImagr', 15],
              1409 : ['eConnectionFail_SinceBeginPageMove', 13], 1410 : ['eConnectionFail_SinceEndPageMove', 15], 1411 : ['eConnectionFail_SinceBeginSave', 14],
              1412 : ['eConnectionFail_SinceEndSave', 15]}

# 표에 실제로 표기될 항목들 [항목이름, 실제 이슈 합계 포함여부, 개수]
status = {1 : ['PC Office 종료 실패', 1], 2 : ['오픈실패', 1], 3 : ['저장실패', 1],
          4 : ['오픈 중 뷰어모드 전환', 1], 5 : ['PC Office 비정상 종료', 1], 6 : ['Crash', 1],
          7 : ['Terminate', 1], 8 : ['(Timeout - validation open)', 0], 9 : ['MS오픈실패', 1],
          10 : ['MS복구팝업', 1], 11 : ['MS제한된보기', 1], 12 : ['(TimeOut - unable total load)', 0],
          13 : ['(TimeOut - unable page move)', 0], 14 : ['(TimeOut - unable save)', 0], 15 : ['(TimeOut - Etc)', 0],
          16 : ['(테스트 프로그램 에러 - 빈 로그 메시지)', 0], 17 : ['(테스트 프로그램 에러 - 알 수 없는 에러)', 0],
          50 : ['기타', 1]}
# 개수 항목 추가
for value in status.values():
    value.append(0)

# msofficevalidation을 하는 확장자
msofficevalidation = ['.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx']

# xml(신버전)
is_xml = ['.docx', '.pptx', '.xlsx']

# text 파일을 줄 별로 자른다.
def LineParsing(line):
    elements = line.split('\t')
    for i in range(len(elements)):
        elements[i] = elements[i].replace("\n", "")
    elements = list(filter(None, elements))

    return elements


def CountErrors(filepath, buildfilepath, ChipSet):
    if os.path.isfile(filepath) and os.path.isfile(buildfilepath):
        f = open(filepath, 'r', encoding='UTF16')
        # 첫 줄은 메타 데이터이므로 넘어감.
        line = f.readline()
        # extname을 담는 리스트
        form = []
        # extname별 에러를 담는 리스트
        error_per_form = []
        # extname count
        count_per_extname = {}
        while line:
            line = f.readline()
            elements = LineParsing(line)

            # [''] 제외
            if elements:
                if elements[0] in count_per_extname.keys():
                    count_per_extname[elements[0]] += 1
                else:
                    count_per_extname[elements[0]] = 1
                # extname이 form에 없을 때
                if elements[0] not in form:
                    # form에 추가
                    form.append(elements[0])
                    # status dict를 copy
                    temp = copy.deepcopy(status)
                    # form의 순서에 맞춰 dict들의 list에 저장
                    error_per_form.append(temp)
                    # errorno
                    temp = int(elements[len(elements)-2])
                    # errorno가 0이 아니면
                    if temp != 0:
                        # 정의된 error이면
                        if temp in errortypes:
                            # Error와 매치되는 항목의 개수를 증가시킨다.
                            error_per_form[len(form)-1][errortypes[temp][1]][2] += 1
                        # 없으면
                        else:
                            # 기타로 분류해서 count를 증가시킴.
                            error_per_form[len(form)-1][50][2] += 1
                    # errorno가 0이면
                    else:
                        # msofficevalidationcheck를 확인
                        temp = int(elements[len(elements)-1])
                        if form[-1] in msofficevalidation:
                            # 1이 아니면
                            if temp != 1:
                                # 정의된 error 이면
                                if temp in errortypes:
                                    # Error와 매치되는 항목의 개수를 증가시킨다.
                                    error_per_form[len(form) - 1][errortypes[temp][1]][2] += 1
                                # 정의된 error가 아니면
                                else:
                                    # 기타에 추가
                                    error_per_form[len(form)-1][50][2] += 1
                        else:
                            # 0이 아니면
                            if temp != 0:
                                # 정의된 error 이면
                                if temp in errortypes:
                                    # Error와 매치되는 항목의 개수를 증가시킨다.
                                    error_per_form[len(form) - 1][errortypes[temp][1]][2] += 1
                                # 정의된 error가 아니면
                                else:
                                    # 기타에 추가
                                    error_per_form[len(form) - 1][50][2] += 1
                # extname이 form에 있을 때
                else:
                    # error_per_form에서 해당하는 index를 찾기 위해 form에서 extname이 있는 index를 찾음.
                    ind = form.index(elements[0])
                    # errorno
                    temp = int(elements[len(elements)-2])
                    if temp != 0:
                        if temp in errortypes:
                            error_per_form[ind][errortypes[temp][1]][2] += 1
                        else:
                            error_per_form[ind][50][2] += 1
                    else:
                        if form[ind] in msofficevalidation:
                            # msofficevalidationcheck
                            temp = int(elements[len(elements)-1])
                            if temp != 1:
                                if temp in errortypes:
                                    error_per_form[ind][errortypes[temp][1]][2] += 1
                                else:
                                    error_per_form[ind][50][2] += 1
                        else:
                            # msofficevalidationcheck
                            temp = int(elements[len(elements) - 1])
                            if temp != 0:
                                if temp in errortypes:
                                    error_per_form[ind][errortypes[temp][1]][2] += 1
                                else:
                                    error_per_form[ind][50][2] += 1
        f.close()

        if '.ppt' in form or '.pptx' in form:
            if len(form) == 1:
                if '.ppt' in form:
                    form.append('.pptx')
                    temp = copy.deepcopy(status)
                    error_per_form.append(temp)
                else:
                    form.append('.ppt')
                    temp = copy.deepcopy(status)
                    error_per_form.append(temp)

        elif '.xls' in form or '.xlsx' in form:
            if len(form) == 1:
                if '.xls' in form:
                    form.append('.xlsx')
                    temp = copy.deepcopy(status)
                    error_per_form.append(temp)
                else:
                    form.append('.xls')
                    temp = copy.deepcopy(status)
                    error_per_form.append(temp)

        # BuildFile Parsing
        buildfileinfo = []
        f = open(buildfilepath, 'r', encoding='UTF8')
        # 파일의 최종 수정 시간을 가져온다.
        last_modify_time = os.path.getmtime(buildfilepath)
        temp = str(datetime.datetime.fromtimestamp(last_modify_time))
        temp = temp.split(' ')
        last_modify_time = '['
        for i in range(2, len(temp[0])):
            if temp[0][i] != '-':
                last_modify_time += temp[0][i]
        last_modify_time += ']'
        buildfileinfo.append(last_modify_time)

        with open(buildfilepath, 'r') as fp:
            for i, line in enumerate(fp):
                if i == 5:
                    # Engine URL 추출
                    ind = line.find('=')
                    buildURL = line[ind+2:len(line)]

                    if 'trunk' in line:
                        buildfileinfo.append('[Trunk]')
                    else:
                        buildfileinfo.append('[Int_G]')
                    break
        buildfileinfo.append('[{}]'.format(ChipSet))

        if len(form) > 1 and form[0] in is_xml:
            temp = form[0]
            form[0] = form[1]
            form[1] = temp

            temp = error_per_form[0]
            error_per_form[0] = error_per_form[1]
            error_per_form[1] = temp

        return [error_per_form, buildfileinfo, form, buildURL, count_per_extname]
    else:
        # 파일이 존재하지 않을 경우 프로그램을 종료한다.
        if os.path.isfile(filepath) == False:
            print("ErrorFile이 존재하지 않습니다.")
            loggingman.write("ErrorFile이 존재하지 않습니다.")
        if os.path.isfile(buildfilepath) == False:
            print("BuildFile이 존재하지 않습니다.")
            loggingman.write("BuildFile이 존재하지 않습니다.")
        quit()
