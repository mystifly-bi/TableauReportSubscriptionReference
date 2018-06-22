__author__ = 'mrinal'
import pandas as pd
import numpy as np
import os, sys, glob
from datetime import datetime, timedelta
from sqlalchemy import create_engine
class Report:
    def getimage(self):

        rep_datetime = datetime.today() - timedelta(1)
        rep_date=rep_datetime.date().strftime("%Y_%b_%d")
        site = 'BI-Reports'

        rpath = 'C:/Users/******/Desktop/Reports/Expedia'
        if os.path.exists(rpath):
            pass
        else:
            os.makedirs(rpath)


        os.chdir('C:\Tabcmd\Command Line Utility')
        download_path1 = rpath + '/' + 'Confirmed Tickets ' + str(rep_date) + '.pdf'
        download_path2 = rpath + '/' + 'Cancelled Tickets ' + str(rep_date) + '.pdf'
        download_path3 = rpath + '/' + 'Exchanged Tickets' + str(rep_date) + '.pdf'
        download_path4 = rpath + '/' + 'Void Tickets  ' + str(rep_date) + '.pdf'
        download_path5 = rpath + '/' + 'Credit Note Generated For ' + str(rep_date) + '.pdf'
        download_path6 = rpath + '/' + 'Searches & Conversions ' + str(rep_date) + '.pdf'
        download_path7 = rpath + '/' + 'Ticket Issuance TAT  ' + str(rep_date) + '.pdf'

        print(download_path1)
        try:
            os.system("tabcmd login -s http://192.168.0.000 -u admin -p 'Password'")
            os.system("tabcmd export -s http://192.168.0.000 -p W@ssPord -t {} \"Expedia/ConfirmedTickets\" --pdf -f \"{}\"".format(site, download_path1))
            os.system("tabcmd export -s http://192.168.0.000 -p W@ssPord -t {} \"Expedia/CancelledTickets\" --pdf -f \"{}\"".format(site, download_path2))
            os.system("tabcmd export -s http://192.168.0.000 -p W@ssPord -t {} \"Expedia/ExchangedTickets\" --pdf -f \"{}\"".format(site, download_path3))
            os.system("tabcmd export -s http://192.168.0.000 -p W@ssPord -t {} \"Expedia/VoidTickets\" --pdf -f \"{}\"".format(site, download_path4))
            os.system("tabcmd export -s http://192.168.0.000 -p W@ssPord -t {} \"Expedia/CreditNoteGeneratedFor\" --pdf -f \"{}\"".format(site, download_path5))
            os.system("tabcmd export -s http://192.168.0.000 -p W@ssPord -t {} \"Expedia/SearchesConversion\" --pdf -f \"{}\"".format(site, download_path6))
            os.system("tabcmd export -s http://192.168.0.000 -p W@ssPord -t {} \"Expedia/TicketIssuanceTAT\" --pdf -f \"{}\"".format(site, download_path7))
            print("Generated report for Expedia :: " )
        except Exception as e:
            print("Failed to generate report for Expedia :: " )
