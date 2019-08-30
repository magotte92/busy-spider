import xlsxwriter
import pandas as pd

counter = 0
# Create a new workbook and add a worksheet
workbook = xlsxwriter.Workbook('hyperlink.xlsx')
worksheet = workbook.add_worksheet('New Speadsheet')
# Format the second column
worksheet.set_column('B:B')
rond = pd.read_csv('last_call.csv', sep=';')


def hyper_maker(db):     # mock link should build down
    global counter
    for i in db.values:
        link = db.loc[counter, 'link']
        name = str(counter)
        mail = db.loc[counter, 'emails']
        try:
            worksheet.write_url('B' + name, link, string='Location of the email in google')
        except Exception as err:
            worksheet.write('B' + name, link.split('/'))
            print(err)
        worksheet.write('A' + name, mail)
        counter += 1


hyper_maker(rond)


workbook.close()