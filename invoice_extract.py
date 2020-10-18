# importing required packages for text extracting
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

import re
from os import listdir
import pgeocode
import locale
from xlwt import Workbook


def convert_pdf_to_txt(path):
    # initalizing empty list
    a = []

    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()

    # Extracting the data from the user
    customer = re.findall("Customer Code[\S\n ]+porter", text)
    amount = re.findall("GRAND TOTAL[\S\n ]+Terms", text)

    customer_content = customer[0]

    list_full = customer_content.split('\n')

    customer_code = re.findall('Customer Code:.*', customer_content)
    a.append(customer_code[0][14:])  # Customer Code

    customer_name = list_full[2]
    a.append(customer_name)  # Customer name

    customer_invoice = re.findall('Invoice Number   :.*', customer_content)
    a.append(customer_invoice[0][19:])  # Invoice number

    customer_tax_invoice = re.findall('Customer Ref No.:.*', customer_content)
    a.append(customer_tax_invoice[0][17:])  # Tax Invoice number

    customer_date = re.findall('Date:.*', customer_content)
    a.append(customer_date[0][5:])  # Invoice date

    postal_code = re.findall('\d{6}[ ]\w\w', customer_content)
    if len(postal_code) != 0:
        postal_code = postal_code[0][:6]
        location_set = pgeocode.Nominatim('in')
        place = location_set.query_postal_code(postal_code)
        a.append(place.county_name)  # State name
        a.append(place.state_name)  # City name
    else:
        a.append('Check')  # City name
        a.append('Check')  # State name

    try:
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

        digit = re.findall("\d*\.\d+|\d+\,\d*\.\d+|\d+", amount[0])
        digit_f = [locale.atof(num) for num in digit]
        position = digit_f.index(max(digit_f))
        a.append(max(digit_f))  # Order quantity
        a.append(digit_f[position - 2])  # Invoice value
    except Exception as e:
        a.append("Check")  # order quantity
        a.append("Check")  # invoice value

    a.append(postal_code)  # Pin Code

    print(f'Invoice {path} successfully written to excel')

    return a


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    li = listdir('/home/kiran/PycharmProjects/pythonProject/File')

    # Creating excel workbook
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')

    for pos, file_name in enumerate(li):
        invoice_values = convert_pdf_to_txt(f'/home/kiran/PycharmProjects/pythonProject/File/{file_name}')
        sheet1.write(pos + 1, 4, invoice_values[0])
        sheet1.write(pos + 1, 5, invoice_values[1])
        sheet1.write(pos + 1, 10, invoice_values[2])
        sheet1.write(pos + 1, 11, invoice_values[3])
        sheet1.write(pos + 1, 12, invoice_values[4])
        sheet1.write(pos + 1, 7, invoice_values[5])
        sheet1.write(pos + 1, 8, invoice_values[6])
        sheet1.write(pos + 1, 15, invoice_values[7])
        sheet1.write(pos + 1, 13, invoice_values[8])
        sheet1.write(pos + 1, 23, invoice_values[9])

    wb.save('invoice.xlsx')
