# # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# The goal of this python file is to get the email        #
# address of activated users.                             #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # #

import xlrd


# Define the function to process the date inside the excel file
def data_processing(filename, sheet_number):

    workbook = xlrd.open_workbook(filename)

    # Get the name of the sheets
    sheet_names = workbook.sheet_names()

    # Change to the sheet you want to use
    working_sheet = workbook.sheet_by_index(sheet_number)

    # Read the 'Name' 'email address' 'Status' row
    for i in range(1):
        Name = working_sheet.col_values(i)[1:]
        email_address = working_sheet.col_values(i + 1)[1:]
        Status = working_sheet.col_values(i + 2)[1:]

    # Check the status of users to see whether to send the email to him/her or not
    index = [i for i, j in enumerate(Status) if j == 'activated']

    # Get the users' email address whose status is activated
    receiver = []
    for i in index:
        receiver.append(email_address[i])

    return sheet_names, receiver
