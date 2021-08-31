#실행 시 옵션을 파싱하고 코드 내의 변수에 값을 넣는다.

import argparse

def parsingargument():
    parser = argparse.ArgumentParser(description="Enter Sender's Email, PW and Receiver's Email. You can also contain CC.")

    parser.add_argument('--ErrorFile', '-EF', help='Path of Error Text File, You must wrap Filepath with "". Ex) "C:\\blahblah\\file.txt"')
    parser.add_argument('--BuildFile', '-BF', help='Path of Build Text File, You must wrap Filepath with "".')
    parser.add_argument('--ChipSet', '-CS', help='Type of ChipSet')
    parser.add_argument('--SenderEmail', '-SE', type=str, action='store', help="Sender's Email")
    parser.add_argument('--Password', '-PW', help="Sender's PW")
    parser.add_argument('--ReceiverEmail', '-RE', help="Receiver's Email")
    parser.add_argument('-CC', help="Carbon Copy", default="")

    args = parser.parse_args()

    return args