# Shift Report Auditing tool for OIT Shift Reports
# Parses all shift reports and returns key information from those which are empty
# Be sure to have ReportTable.csv in the same folder as this program
# ⤷ Go to Sharepoint > Site Contents, click 'Shift Reports' List
# ⤷ Click 'Export to Excel' and save as 'ReportTable.csv'

import csv

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

#opening file...
file = open('ReportTable.csv')
type(file)
csvreader = csv.reader(file)
next(csvreader)

threshold = input("Welcome to the OIT Shift Report Audit Tool. Set threshold for logs per hour: ")
#create lists of shiftID, netID, shift location, and shift date and time for each shift
ids = []
netids = []
locs = []
dates = []
lph = [] #logs per hour

#iterate through file
prevTimeStr = ""
for row in csvreader:
        startTime = row[4]
        endTime = row[6] 

        shiftLength = calculateTime(startTime, endTime) #calculates length of shift
        
        endTimeStr = " -", endTime[endTime.index(' '):] #shift end time
        timeStr = row[4] + ''.join(endTimeStr) #concat with date and start time
        
        logsPerHr = int(row[9])/shiftLength #total logs recorded divided by shift length

        if logsPerHr <= int(threshold) and timeStr != prevTimeStr: #if less than 2 logs per hour and not a duplicate report
            ids.append(row[0]) #shift ID
            locs.append(row[2]) #location
            netids.append(row[5]) #netID
            dates.append(timeStr) #date and time
            lph.append(str(float(round(logsPerHr,3))))
            prevTimeStr = timeStr 

#writing data to text file
f = open("output.txt", "a")
f.truncate(0) #clear any data from file
output = []

#iterate through each list and add data to file
for i in range(len(ids)):
    vals = "Shift ID: ", ids[i], "\nNetID: ", netids[i], "\nLocation: ", locs[i], "\nDate: ", dates[i],"\nLogs Per Hour: ", lph[i]
    text = ''.join(vals)
    f.write(text)
    f.write("\n-------------------------------------\n")
f.close()

#writing data to csv file
with open('output.csv', mode='w') as csv_file:
    fieldnames = ['id', 'netid', 'loc', 'date']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

    writer.writeheader()
    for i in range(len(ids)):
        writer.writerow({'id': ids[i], 'netid': netids[i], 'loc': locs[i], 'date' : dates[i]})

file.close()