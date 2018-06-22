__author__ = 'mrinal'
from datetime import datetime, timedelta
import sys, os
import pdfGeneration, emailReport
if __name__ == "__main__":
    print("Report Generation Started at : ", str(datetime.now()))

    obj1 = pdfGeneration.Report()
    obj1.getimage()

    obj2 = emailReport.EmailSend()
    obj2.sendmail()


    print("Report Generation completed at : ", str(datetime.now()))
