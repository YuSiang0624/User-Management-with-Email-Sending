# # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# The goal of this python file is to send the message to  #
# the email address which come from Data_Processing.py    #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # #

from Data_Processing import data_processing
import yagmail


def email_sending(subject, contents, sheet_number, attachment=None):

    sheet_names, receiver = data_processing('Data.xlsx', sheet_number)

    # Connect to smtp server
    yag_smtp_connection = yagmail.SMTP(user='xxx@gmail.com', password='0000', host='smtp.gmail.com')

    # send the email
    yag_smtp_connection.send(receiver, subject, contents, attachment)
