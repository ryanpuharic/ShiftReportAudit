# Shift Report Auditing tool for OIT Shift Reports
# Parses all shift reports and returns key information from those which are empty
# Be sure to have ReportTable.csv in the same folder as this program
# ⤷ Go to Sharepoint > Site Contents, click 'Shift Reports' List
# ⤷ Click 'Export to Excel' and save as 'ReportTable.csv'

import csv
import datetime

#opening reportTable...
reportTable = open('ReportTable.csv')
type(reportTable)
reportReader = csv.reader(reportTable)
next(reportReader)

#opening shiftTable...
shiftTable = open('ShiftTable.csv')
type(shiftTable)
shiftReader = csv.reader(shiftTable)
next(shiftReader)

#set interval of dates
interval = ''
while(len(interval)!= 21 or '-' not in interval):
    interval = input("Welcome to the OIT Shift Report Audit Tool. \nLast Update: 12/13/2021 by Ryan Puharic. \nPlease enter the interval of dates you would like to audit between in this format: \nmm/dd/yyyy-mm/dd/yyyy \n\n")
    print()
    
startDate = interval[:interval.index('-')]
endDate = interval[interval.index('-')+1:]
d1 = datetime.date(int(startDate[6:10]), int(startDate[:2]), int(startDate[3:5]))
d2 = datetime.date(int(endDate[6:10]), int(endDate[:2]), int(endDate[3:5]))

#begin audit
print("Auditing...")

#create lists of shiftID, netID, shift location, and shift date and time for each shift
ids = []
netids = []
locs = []
dates = []
allDates = []

#check if empty report
for row in reportReader:
    startTime = row[4]
    endTime = row[6] 
    endTimeStr = " -", endTime[endTime.index(' '):] #shift end time
    timeStr = row[4] + ''.join(endTimeStr) #concat with date and start time
    allDates.append(timeStr)

    justDate = startTime[:startTime.index(' ')]

    #inserting 0 before single digits (ex: 01 instead of 1)
    if(justDate[1] == '/'):
        justDate = '0' + justDate
    if(justDate[4] == '/'):
        justDate = justDate[:3] +'0' + justDate[3:]

    d3 = datetime.date(int(justDate[6:10]), int(justDate[:2]), int(justDate[3:5]))

    if(d3 >= d1 and d3 <= d2): #if date is in desired interval
        if (int(row[9]) == 0 and timeStr not in dates): #if empty report and not a duplicate report 
            ids.append(row[0]) #shift ID
            locs.append(row[2]) #location
            netids.append(row[5]) #netID
            dates.append(timeStr) #date and time

#check if no report
nrDates = [] #date and time of shifts with no reports
nrLocs = []

for s in shiftReader:
    #string manipulation because ShiftTable and ReportTable dates and times are formatted differently
    shiftDate = s[3]
    shiftTime = s[5]
    shiftTimeStr = shiftDate[:shiftDate.index(' ')-2] + "20" + shiftDate[shiftDate.index(' ')-2:shiftDate.index(' ')] + " " #will not work after the year 2099

    #inserting 0 before single digits (ex: 01 instead of 1)
    if(shiftTimeStr[1] == '/'):
        shiftTimeStr = '0' + shiftTimeStr
    if(shiftTimeStr[4] == '/'):
        shiftTimeStr = shiftTimeStr[:3] +'0' + shiftTimeStr[3:]
    
    #makes sure years are in yyyy format
    if(len(shiftTimeStr) == 13):
        shiftTimeStr = shiftTimeStr[:6] +shiftTimeStr[8:]

    d3 = datetime.date(int(shiftTimeStr[6:10]), int(shiftTimeStr[:2]), int(shiftTimeStr[3:5]))

    if(d3 >= d1 and d3 <= d2): #if date is in desired interval
        if(shiftTimeStr[3] == '0'):
            shiftTimeStr = shiftTimeStr[:3] + shiftTimeStr[4:]
        if(shiftTime[0] == '0'):
            shiftTime = shiftTime[1:]
        if(shiftTime[shiftTime.index('-')+2] == '0'):
            shiftTime = shiftTime[:shiftTime.index('-')+2] + shiftTime[shiftTime.index('-')+3:]
        shiftTimeStr = shiftTimeStr + shiftTime

        #check if report shift data matches to any done report
        noReport = True
        for d in allDates:
            if(shiftTimeStr == d):
                noReport = False
        if(noReport and 'Sups' not in s[1]):
            if(shiftTimeStr not in nrDates):
                nrDates.append(shiftTimeStr)
                nrLocs.append(s[1])

#write data to csv files
with open('empty.csv', mode='w') as csv_file:
    fieldnames = ['id', 'netid', 'loc', 'date']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

    writer.writeheader()
    for i in range(len(ids)):
        writer.writerow({'id': ids[i], 'netid': netids[i], 'loc': locs[i], 'date' : dates[i]})

with open('noreports.csv', mode='w') as csv_file:
    fieldnames = ['date', 'loc']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

    writer.writeheader()
    for i in range(len(nrDates)):
        writer.writerow({'date' : nrDates[i], 'loc' : nrLocs[i]})

reportTable.close()
print('Audit Complete.')

'''
OBSELETE STUFF
USED FOR SHIFT REPORT THRESHOLD

lph = [] #logs per hour
shiftLength = calculateTime(startTime, endTime) #calculates length of shift
logsPerHr = int(row[9])/shiftLength #total logs recorded divided by shift length


#writing data to text file
f = open("output.txt", "a")
f.truncate(0) #clear any data from file
output = []
lph.append(str(float(round(logsPerHr,3))))

#iterate through each list and add data to file
for i in range(len(ids)):
    vals = "Shift ID: ", ids[i], "\nNetID: ", netids[i], "\nLocation: ", locs[i], "\nDate: ", dates[i],"\nLogs Per Hour: ", lph[i]
    text = ''.join(vals)
    f.write(text)
    f.write("\n-------------------------------------\n")
f.close()

#calculates length of shifts
def calculateTime(startTime, endTime):
    hr = int(endTime[endTime.index(' '):endTime.index(':')]) - int(startTime[startTime.index(' '):startTime.index(':')]) #difference between hours worked
    if(hr < 0): #if shift switches from AM to PM
        hr = 12 + hr

    mi = int(endTime[endTime.index(':')+1:endTime.index('M')-2]) - int(startTime[startTime.index(':')+1:startTime.index('M')-2]) #difference between minutes worked
    if(mi < 0): #adjusts for when shifts go from :45 - > :00 or :30 -> :00
        mi = 60 + mi
        hr = hr - 1
    return hr + mi/60
'''