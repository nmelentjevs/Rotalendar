import xlrd
from xlutils.copy import copy
import datetime
import dateutil.parser as parser
try:
    # for Python2
    from Tkinter import *
except ImportError:
    # for Python3
    from tkinter import *

from tkinter.filedialog import askopenfilename

def pressed():
    global name
    name = e1.get()
    global surname
    surname = e2.get()
    root.quit()


root = Tk()


Label(root, text="First Name").grid(row=0)
Label(root, text="Last Name").grid(row=1)

e1 = Entry(root)
e2 = Entry(root)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)


Button(root, text='Ok', command=pressed).grid(row=3, column=0, sticky=W, pady=4)


mainloop()

root.withdraw()
root.update()
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
root.update()
root.destroy()
print(filename)


wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_index(0)

properdates = []


def finddates():
    global numdate
    numdate = []
    for i in range(sheet.ncols):
        try:
            numdate.append(int(sheet.cell(1, i).value))
        except ValueError:
            continue
    for i in numdate:
        properdates.append(xlrd.xldate.xldate_as_datetime(i, wb.datemode).strftime('%m/%d/%Y'))
    return properdates


def findrowbyname(name):
    for row in range(sheet.nrows):
        for column in range(sheet.ncols):
            if name == sheet.cell(row, column).value:
                return row


def findshiftsbyrow(row):
    shifts = []
    for i in range(sheet.ncols):
        shifts.append(sheet.cell(row, i).value)
    return shifts[2:9]


shiftvalues = {"6p6a": "18:00-6:00", "LN": "20:00-6:00", "GN": "8:00-16:00", "G": "6:00-14:00", "8": "Day Off", "N": "22:00-6:00", "M": "20:00-4:00", "6p2a": "18:00-2:00", "D": "14:00-22:00", "ED": "12:00-20:00", "EM": "10:00-20:00", "10a6p": "10:00-18:00", "_": "Day Off"}

shiftnames = ["LN", "6p6a", "N", "M", "D", "ED", "6p2a", "GN", "G", "8"]

transformshifts = []

for i in findshiftsbyrow(findrowbyname(surname + ', ' + name)):
    transformshifts.append(shiftvalues.get(i))

finddates()

daysandshitfs = dict(zip(properdates, transformshifts))

calendardatesstart = []
calendardatesfinish = []


for i, k in zip(range(len(properdates)), properdates):
    if daysandshitfs[k] != 'Day Off':
        date = parser.parse(properdates[i])
        calendardatesstart.append(date.isoformat()[:11] + daysandshitfs[k][:-5]+':00')

for i, k in zip(range(len(properdates)), properdates):
    try:
        if daysandshitfs[k] == '14:00-22:00' or daysandshitfs[k] == '12:00-20:00':
            date = parser.parse(properdates[i])
            calendardatesfinish.append(date.isoformat()[:11] + daysandshitfs[k][6:10] + ':00')
        elif daysandshitfs[k] != 'Day Off':
            date = parser.parse(properdates[i+1])
            calendardatesfinish.append(date.isoformat()[:11] + daysandshitfs[k][6:10]+':00')
    except IndexError:
        numdate.append(numdate[-1] +1)
        properdates.append(xlrd.xldate.xldate_as_datetime(numdate[-1], wb.datemode).strftime('%m/%d/%Y'))
        date = parser.parse(properdates[i])
        calendardatesfinish.append(date.isoformat()[:11] + daysandshitfs[k][6:10] + ':00')


print(calendardatesstart)
print(calendardatesfinish)

