#TODO handle negative (error) values when computing aggregates in GaugeDataset and TSYear classes

import uno
import datetime
from calendar import monthrange

from com.sun.star.table.CellContentType import EMPTY, VALUE, TEXT, FORMULA

def Main():

    document = XSCRIPTCONTEXT.getDocument()
    #sheet = document.getSheets().getByIndex(0) #gets first sheet
    sheet = document.CurrentController.ActiveSheet #gets current open sheet when macro is ran

    dataSet = GaugeDataset(1)

    #datastructure building loop
    for row in range (1, 10248575):
    #for row in range (1, 3000):
        if sheet.getCellByPosition(0, row).Type == EMPTY:
            break
        
        if (sheet.getCellByPosition(1, row).Type != EMPTY):
            calcDate = sheet.getCellByPosition(0, row).Value
            date = datetime.datetime(1899, 12, 30)
            date += datetime.timedelta(days = calcDate)
            dataSet.AddRecord(date.year, date.month, date.day, sheet.getCellByPosition(1, row).Value)
            

    # ##test 
    # sheet.getCellByPosition(2, 0).Value = sheet.getCellByPosition(0, 1).Value
    # sheet.getCellByPosition(2, 1).Value = calcDate
    # sheet.getCellByPosition(2, 2).Value = date.year
    # sheet.getCellByPosition(2, 3).Value = date.month
    # sheet.getCellByPosition(2, 4).Value = date.day
    # sheet.getCellByPosition(2, 6).Value = dataSet.GetRecordsYearCount()
    # sheet.getCellByPosition(2, 7).Value = dataSet.minYear
    # sheet.getCellByPosition(2, 8).Value = dataSet.maxYear
    # return
    # ##end test
 
    #output writing loop

    #headers
    #todo move this to own function (probably should refactor this to have writing to cells in its own object/function)
    sheet.getCellByPosition(3,0).String = "Year"
    sheet.getCellByPosition(4,0).String = "Month"
    sheet.getCellByPosition(5,0).String = "Total Rainfall"
    sheet.getCellByPosition(6,0).String = "Max Daily Rainfall"
    sheet.getCellByPosition(7,0).String = "Average Daily Rainfall"
    sheet.getCellByPosition(8,0).String = "Rainy Days"
    sheet.getCellByPosition(9,0).String = "Missing Records"
    
    sheet.getCellByPosition(11,0).String = "Year"
    sheet.getCellByPosition(12,0).String = "Total Rainfall"
    sheet.getCellByPosition(13,0).String = "Max Daily Rainfall"
    sheet.getCellByPosition(14,0).String = "Rainy Days"
    sheet.getCellByPosition(15,0).String = "Missing Records"
    
    sheet.getCellByPosition(17,0).String = "Month"
    sheet.getCellByPosition(18,0).String = "Average Monthly Rainfall"
    sheet.getCellByPosition(19,0).String = "Average Max Daily Rainfall"

    #First data output loop
    for row in range (1, dataSet.GetRecordsYearCount() + 2):
        currentYear = dataSet.minYear + row - 1
        #TODO consider merging the cell with year written bellow with 11 other bellow it.
        sheet.getCellByPosition(3, 1 + ((row - 1) * 12)).Value = currentYear
        sheet.getCellByPosition(11, row).Value = currentYear

        #in case an entire year had missing records. There would be no entry in the dataset dictionary
        if (not(currentYear in dataSet.records)):
            continue

        #inner loop for monthly data
        for month in range (1, 13):
            sheet.getCellByPosition(4, month + (12 * (row - 1))).Value = month
            sheet.getCellByPosition(5, month + (12 * (row - 1))).Value = dataSet.records[currentYear].GetTotalRainfallMonth(month)
            sheet.getCellByPosition(6, month + (12 * (row - 1))).Value = dataSet.records[currentYear].GetMaxDailyMonth(month)
            sheet.getCellByPosition(7, month + (12 * (row - 1))).Value = dataSet.records[currentYear].GetAverageRainMonth(month)
            sheet.getCellByPosition(8, month + (12 * (row - 1))).Value = dataSet.records[currentYear].GetRainyDaysMonth(month)
            sheet.getCellByPosition(9, month + (12 * (row - 1))).Value = dataSet.records[currentYear].GetMissingRecordsMonth(month)
        
        #outerloop for annual data
        sheet.getCellByPosition(12, row).Value = dataSet.records[currentYear].GetTotalRainfallAnnum()
        sheet.getCellByPosition(13, row).Value = dataSet.records[currentYear].GetMaxDailyAnnum()
        sheet.getCellByPosition(14, row).Value = dataSet.records[currentYear].GetRainyDaysAnnum()
        sheet.getCellByPosition(15, row).Value = dataSet.records[currentYear].GetMissingRecordsAnnum()

    
    #This loop is for aggregated monthly data
    for month in range (1, 13):
        sheet.getCellByPosition(17, month).Value = month
        sheet.getCellByPosition(18, month).Value = dataSet.GetAverageMonthlyRainfall(month)
        sheet.getCellByPosition(19, month).Value = dataSet.GetAverageMaxDailyRainfallMonth(month)

        
class GaugeDataset:
    def __init__(self, _rainThreshold = 1):
        self.records = {}
        self.minYear : int = 3000
        self.maxYear : int = 1000
        self.rainThreshold : int = 1 # rainThreshold is minimum rainfall reading to count as a rainy day.
        self.rainThreshold = _rainThreshold
        return
    
    def AddRecord(self, year : int, month : int, day : int, rainfall):
        if not(year in self.records):
            self.records[year] = TSYear(year, self.rainThreshold)
            self.minYear = min(year, self.minYear)
            self.maxYear = max(year, self.maxYear)

        self.records[year].AddRecord(month, day, rainfall)        

    def GetRecordsYearCount(self):
        return len(self.records)

    def GetAverageMonthlyRainfall(self, month : int): #for entire time series.
        average : float = 0
        counter : int = 0

        for key, value in self.records.items():
            average += value.GetTotalRainfallMonth(month)
            counter += 1

        return average / counter
    
    def GetAverageMaxDailyRainfallMonth(self, month : int):
        average : float = 0
        counter : int = 0

        for key, value in self.records.items():
            average += value.GetMaxDailyMonth(month)
            counter += 1

        return average / counter

    
#TODO add a conditional based on missing records. Set a threshold for missing records over which the month (or year) is excluded from averages/totals.
class TSYear:
    #For recordExtra: key is month, value is a array of three values: [0] missing records, [1] rainy days, [2] max daily, [3] cummulative daily
    #cummulative daily is also used to compute average day per month

    def __init__(self, _year : int, _rainThreshold = 1):
        #python idiocy. Putting these variables outside constructor makes them "universal" to all class instances.
        self.recordExtra = {}
        self.year = _year
        self.rainThreshold = _rainThreshold
        for month in range(1, 13):
            self.recordExtra[month] = [monthrange(self.year, month)[1], 0, -1,  0]

    def AddRecord(self, month : int, day : int, rainfall):
        self.recordExtra[month][0] -= 1
        if (rainfall >= self.rainThreshold): self.recordExtra[month][1] += 1
        if (rainfall > self.recordExtra[month][2]) : self.recordExtra[month][2] = rainfall
        self.recordExtra[month][3] += rainfall

    def GetMissingRecordsMonth(self, month : int):
        return self.recordExtra[month][0]
    
    def GetMissingRecordsAnnum(self):
        totalMissing :int = 0
        for month in range(1, 13):
            totalMissing += self.GetMissingRecordsMonth(month)
        return totalMissing
    
    def GetRainyDaysMonth(self, month : int):
        return self.recordExtra[month][1]
    
    def GetRainyDaysAnnum(self):
        totalDays :int = 0
        for month in range(1, 13):
            totalDays += self.GetRainyDaysMonth(month)
        return totalDays

    def GetMaxDailyMonth(self, month : int):
        return self.recordExtra[month][2]

    def GetMaxDailyAnnum(self):
        maxVal : int = -1
        for month in range(1, 13):
            maxVal = max(self.GetMaxDailyMonth(month), maxVal)
        return maxVal
    
    def GetTotalRainfallMonth(self, month : int):
        return self.recordExtra[month][3]
    
    def GetTotalRainfallAnnum(self):
        totalRain : int = 0
        for month in range(1, 13):
            totalRain += self.GetTotalRainfallMonth(month)
        return totalRain
    
    def GetAverageRainMonth(self, month : int):
        recordedDays = monthrange(self.year, month)[1] - self.recordExtra[month][0]
        if (recordedDays > 0):
            return self.recordExtra[month][3] / (monthrange(self.year, month)[1] - self.recordExtra[month][0])
        else:
            return -1