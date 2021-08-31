import content
import parsing
import sendmail
import argumentparser
import log


if __name__ == '__main__':
    loggingman = log.loggingman()
    args = argumentparser.parsingargument()
    key = 0
    missingargs = []

    if args.ErrorFile == None:
        key = 1
        missingargs.append('ErrorFile')
    if args.BuildFile == None:
        key = 1
        missingargs.append('BuildFile')
    if args.ChipSet == None:
        key = 1
        missingargs.append('ChipSet')
    # if args.SenderEmail == None:
    #     key = 1
    #     missingargs.append('Sender Email')
    # if args.Password == None:
    #     key = 1
    #     missingargs.append('Password')
    if args.ReceiverEmail == None:
        key = 1
        missingargs.append('ReceiverEmail')

    if key == 0:
        # 기능 실행
        content = content.ProduceContents(parsing.CountErrors(args.ErrorFile, args.BuildFile, args.ChipSet))
        sm = sendmail.send_mail(args.SenderEmail, args.Password)
        sm.SendMail(args.ReceiverEmail, args.CC, content[0], content[1])
    else:
        print('인자 {}를 포함하지 않았습니다.'.format(missingargs))
        loggingman.write('인자 {}를 포함하지 않았습니다.'.format(missingargs))
