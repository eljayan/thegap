import datetime
import openpyxl
from pprint import pprint

#read the excel file

def main(startHour):
    UTCHours = []
    GapDays = []
    FilledDays = []

    wb = openpyxl.load_workbook("D:/myScripts/theGap/table/THEGAP.xlsx")
    sh = wb.active

    for r in sh.rows[1:]:
        hourItem = Hour()
        hourItem.date = ordinalToDate(r[0].value)
        hourItem.open = float(r[1].value)
        hourItem.high = float(r[2].value)
        hourItem.low = float(r[3].value)
        hourItem.close = float(r[4].value)
        UTCHours.append(hourItem)

    # print vars(UTCHours[-1])

    # primero llenar las horas en una lista


    for position, hour in enumerate(UTCHours):
        if hour.date.hour == startHour:
            startPosition = position
            # if startPosition > 15:
            if startPosition > 11:
                break

    #agrupo los elements en gap days
    while startPosition < len(UTCHours):
        gapDay =  Gap()
        gapDay.date = UTCHours[startPosition].date
        # gapDay.elements = (UTCHours[startPosition: startPosition+8])
        gapDay.elements = (UTCHours[startPosition: startPosition+12])
        # gapDay.previousCloseDate = UTCHours[startPosition-16].date
        gapDay.previousCloseDate = UTCHours[startPosition-12].date
        # gapDay.previousClosePrice = UTCHours[startPosition-16].close
        gapDay.previousClosePrice = UTCHours[startPosition-12].close
        GapDays.append(gapDay)
        startPosition += 24


    for gapDay in GapDays:
        gapDay.setPriceHigh()
        gapDay.setPriceLow()
        gapDay.gapFilled()
        if gapDay.filled:
            FilledDays.append(gapDay)

        # print vars(gapDay)

    # pprint(vars(GapDays[0]))
    # pprint(vars(GapDays[1]))
    # pprint(vars(GapDays[2]))
    #crear el objeto gap con las horas especificadas y agragarlo a GapDays

    #ir por cada gap object viendo en cuantas se cerro exitosamente. Y dividirlo para el numero de gap days

    totalDays = len(GapDays)
    filledDays = len(FilledDays)

    probability = float(filledDays)/float(totalDays)

    return "Gap Close Probability: " + str(probability)


class Gap:
    def __init__(self):
        self.elements = []
        self.date = None

    def setPriceHigh(self):
        #iterete throu gapHours looking for the hihgest price
        maxPrice = 0
        for i in self.elements:
            if i.high > maxPrice:
                maxPrice = i.high

        self.priceHigh = maxPrice

    def setPriceLow(self):
        #iterato through gapHours looking for the lowest price
        minPrice = None
        for i in self.elements:
            if minPrice:
                if i.low < minPrice:
                    minPrice = i.low
            else:
                minPrice = i.low

        self.priceLow = minPrice


    def gapFilled(self):
        self.filled =  self.priceHigh > self.previousClosePrice > self.priceLow


def ordinalToDate(ordinalDouble):
    seconds = (ordinalDouble - 25569) * 86400.0
    return datetime.datetime.utcfromtimestamp(seconds)


class Hour:
    def __init__(self):
        self.date = None
        self.open = None
        self.high = None
        self.low = None
        self.close = None




if __name__ == "__main__":
    for r in range(24):
        print r, main(r)