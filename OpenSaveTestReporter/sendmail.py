#Content를 포함해 Email을 보낸다.

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class send_mail:
      # 세션 생성
      # s = smtplib.SMTP('polarisoffice-com.mail.protection.outlook.com', 25)
      s = smtplib.SMTP('smtp.gmail.com', 587)
      def __init__(self, ID, PW):
            # TLS 보안 시작
            send_mail.s.starttls()
            # self.ID = 'svnrndadmin@infraware.co.kr'
            self.ID = ID
            self.PW = PW
            # 로그인 인증
            print(send_mail.s.login(self.ID, self.PW))

      def SendMail(self, To, Cc, text, Subject):
            msg = MIMEMultipart('TEST')
            msg.attach(MIMEText(text, 'html'))
            msg['Subject'] = Subject
            msg['From'] = self.ID
            # To = To.split(',')
            msg['to'] = To
            msg['Cc'] = Cc
            rcvs = Cc.split(',') + To.split(',')
            # 메일 보내기
            print(To)
            print(rcvs)
            send_mail.s.sendmail(self.ID, rcvs, msg.as_string())
            # 세션 종료
            send_mail.s.quit()

