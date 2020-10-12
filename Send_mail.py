import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import datetime, pandas as pd
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')

class runMail(object):
    def __init__(self):
        self.fileread()


    def fileread(self):
        """file reading.... """
        df = pd.read_excel('Excel File')
        uniqueMail = set()
        df = df.astype('str')
        mails = df['Email Addr'].astype('str')

        """ Prepare unique email list """
        for item in mails.tolist():
            uniqueMail.add(item)

        for i in uniqueMail:
            newdf = df[df['Email Addr'] == str(i)]
            downloadData = []
            for r in range(len(newdf)):
                template = """
                              <tr>
                              <td>""" + str(newdf['NAME'].values[r]) + """</td>
                              <td>""" + str(newdf['Description'].values[r]) + """</td>
                              <td>""" + str(newdf['Number'].values[r]) + """</td>
                              </tr>
                              """
                downloadData.append(template)
            downloadData = "\n\r".join(downloadData)
            self.mail_send(str(i), downloadData,str(newdf['NAME'].values[0]))

    def mail_send(self,mails, downloadData,companyName):
        print('\n\tSending Mail........')
        msg = MIMEMultipart() 
        #from_mail = "shrirang.kulkarni@alignedautomation.com"
        #from_mail = outlook.CreateItem(0)
        msg['Subject'] = "Subject"+companyName
        #msg['From'] = from_mail
        msg['To'] = ""

        html = 'BODY'
        msgText = MIMEText(html, 'html')
        msg.attach(msgText)
        s = smtplib.SMTP('smtp.office365.com', 25)
        s.ehlo()
        s.starttls()
        s.ehlo()
        text = msg.as_string()
        #s.sendmail(from_mail, str(mails), text)
        s.quit()
        #s = outlook.CreateItem(0)
        text = msg.as_string()
        s.sendmail(str(mails), text)
        s.Send()

        print('\n\tMail Sent Successfully.')


if __name__ == '__main__':
    runMail()
