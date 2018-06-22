__author__ = 'mrinal'
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import os, smtplib
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import re

class EmailSend:
    def sendmail(self):
        rep_datetime = datetime.today() - timedelta(1)
        rep_date=rep_datetime.date().strftime("%Y_%b_%d")
        # Define these once; use them twice!
        strFrom ='xyz@mystifly.com'
        strTo = "'abc@mystifly.com','def@mystifly.com'"

        os.chdir('C:/Users/*****/Desktop/Reports/Expedia')

        recipients = ['abc@mystifly.com','def@mystifly.com']
        emaillist = [elem.strip().split(',') for elem in recipients]
        msg = MIMEMultipart()

        msg['Subject'] = 'Auto Generated Expedia Report ' +  str(rep_date)
        msg['From'] = 'xyz@mystifly.com'
        msg['To'] = "'abc@mystifly.com','def@mystifly.com'"
        msg.preamble = 'Multipart message.\n'
        reportFile = 'Expedia  ' + str(rep_date) + '.pdf'

        reportFile1 = 'Confirmed Tickets ' + str(rep_date) + '.pdf'
        reportFile2 = 'Cancelled Tickets ' + str(rep_date) + '.pdf'
        reportFile3 = 'Exchanged Tickets' + str(rep_date) + '.pdf'
        reportFile4 = 'Void Tickets  ' + str(rep_date) + '.pdf'
        reportFile5 = 'Credit Note Generated For ' + str(rep_date) + '.pdf'
        reportFile6 = 'Searches & Conversions ' + str(rep_date) + '.pdf'
        reportFile7 = 'Ticket Issuance TAT  ' + str(rep_date) + '.pdf'

        from PyPDF2 import PdfFileWriter, PdfFileReader

        def append_pdf(input,output):
            [output.addPage(input.getPage(page_num)) for page_num in range(input.numPages)]

        output = PdfFileWriter()

        append_pdf(PdfFileReader(open(reportFile1,"rb")),output)
        append_pdf(PdfFileReader(open(reportFile2,"rb")),output)
        append_pdf(PdfFileReader(open(reportFile3,"rb")),output)
        append_pdf(PdfFileReader(open(reportFile4,"rb")),output)
        append_pdf(PdfFileReader(open(reportFile5,"rb")),output)
        append_pdf(PdfFileReader(open(reportFile6,"rb")),output)
        append_pdf(PdfFileReader(open(reportFile7,"rb")),output)


        report_path = 'C:/Users/****/Desktop/Reports/Expedia' + '/' + reportFile
        output.write(open(report_path,"wb"))
        report_path1 = 'C:/Users/*****/Desktop/Reports/Expedia' + '/' + reportFile1
        report_path2 = 'C:/Users/*****/Desktop/Reports/Expedia' + '/' + reportFile2
        report_path3 = 'C:/Users/*****/Desktop/Reports/Expedia' + '/' + reportFile3

        report_path4 = 'C:/Users/*****/Desktop/Reports/Expedia' + '/' + reportFile5
        report_path4 = 'C:/Users/******/Desktop/Reports/Expedia' + '/' + reportFile6
        report_path4 = 'C:/Users/*****/Desktop/Reports/Expedia' + '/' + reportFile7

        if os.path.exists(report_path1):


            body ='''
Hi,

Please find Expedia Report.


Thanks and Regards,
Mrinal Kishore | Data Engineer
E: *******@mystifly.com



            '''
            part = MIMEText(body)
            msg.attach(part)

            # part = MIMEApplication(open(str(report_path1),"rb").read())
            # part.add_header('Content-Disposition', 'attachment', filename=str(reportFile1))
            # msg.attach(part)
            #
            # part = MIMEApplication(open(str(report_path2),"rb").read())
            # part.add_header('Content-Disposition', 'attachment', filename=str(reportFile2))
            # msg.attach(part)
            #
            # part = MIMEApplication(open(str(report_path3),"rb").read())
            # part.add_header('Content-Disposition', 'attachment', filename=str(reportFile3))
            # msg.attach(part)
            #
            # part = MIMEApplication(open(str(report_path4),"rb").read())
            # part.add_header('Content-Disposition', 'attachment', filename=str(reportFile4))
            # msg.attach(part)

            part = MIMEApplication(open(str(report_path),"rb").read())
            part.add_header('Content-Disposition', 'attachment', filename=str(reportFile))
            msg.attach(part)
        else:
            body = '''

Hi,

No Report has been generated.


Thanks and Regards,
Mrinal Kishore | Data Engineer
E: mrinal@mystifly.com


            '''
            part = MIMEText(body)
            msg.attach(part)


        try:
            server = smtplib.SMTP("outlook.office365.com", 587, timeout=120)
            server.ehlo()
            server.starttls()
            server.login("*********@mystifly.com", "Password")
            server.sendmail(msg['From'], emaillist , msg.as_string())
            print("Mail Sent Successfully :: ")
        except Exception as e:
            print("Problem sending mail : ", e)
