# import PyPDF2
import csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
# from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
# import BeautifulSoup
# with open('ProximaNova-Semibold.otf', 'r') as ttfFile:
#    pdfmetrics.registerFont(TTFont("", ttfFile))

'''pdfFileObj = open('17-035.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
pdfReader.numPages
pageObj = pdfReader.getPage(0)
print pageObj.extractText()'''

pdf_file_name = 'HRRG' + ' ' + 'Membership Listing' + '.pdf'
c = canvas.Canvas(pdf_file_name, pagesize = letter)

doc = SimpleDocTemplate("simple_table.pdf", pagesize=letter)
elements = []

data_list = []

with open('test.csv', 'rb') as fin:
    data = csv.reader(fin)
    next(fin)
    for row in data:
        # temp = list(row)
        # fmt =u'{:<15}'*len(temp)
        # print fmt.format(*[s.decode('utf-8') for s in temp])
        account_name = row[0]
        salutation = row[1]
        last_name = row[2]
        first_name = row[3]
        title = row[6]
        data_list.append([salutation + first_name + ' ' + last_name, title, account_name])

t = Table(data_list)
# print t
# t.setStyle(TableStyle([('BACKGROUND', (1, 1), (-2, -2), colors.green),
#                      ('TEXTCOLOR', (0, 0), (1, -1), colors.red)]))
elements.append(t)
# print t
# print elements
doc.build(elements)

#print c.getAvailableFonts()
# Header Text
'''c.setFont('Courier', 18, leading = None)
c.drawRightString(8.25*inch, 10.25*inch, "Human Resources Roundtable Group")
c.setFont('Courier', 10, leading = None)
c.drawCentredString(8*inch, 10*inch, "Brought to you by")
c.setFont('Courier', 10, leading = None)
c.drawCentredString(8.5*inch, 9.5*inch, "Executive Networks, Inc.")
c.setFont('Courier', 14, leading = None)
c.drawCentredString(4.5*inch, 7.5*inch, "Membership List")
# Listing
c.setFont('Courier', 11, leading = None)
c.drawRightString(1*inch, 4*inch, salutation + first_name + ' ' +last_name)
c.setFont('Courier', 10, leading = None)
c.drawRightString(2*inch, 3*inch, title)
c.setFont('Courier', 18, leading = None)
c.drawRightString(4.5*inch, 7.5*inch, account_name)

logo = 'ENlogo.jpg'
c.drawImage(logo, 0.75*inch, 0.5*inch, width=7.25*inch, height=0.84*inch)
c.showPage()

print 'writing'
c.save()'''
