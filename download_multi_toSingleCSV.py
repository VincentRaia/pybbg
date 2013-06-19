from win32com.client import DispatchWithEvents
from pythoncom import PumpWaitingMessages, Empty, Missing
from time import time
from datetime import datetime
import os, csv

list = ["IBM US Equity", "AAPL US Equity"]


entity = 'PX_LAST'
start = datetime(2011, 1, 1)
end = datetime(2012, 12, 31)
bars = 250
csv_file = 'same_test.csv'

def clearCSV(csv_file):
    csv_file = open(csv_file, "w")
    csv_file.truncate()
    writer = csv.writer(open('same_test.csv', 'ab'), delimiter = ',')
    header = ['DATE', 'AAPL_Close', 'IBM_Close']
    writer.writerow(header)
    csv_file.close()
    return

class get_historical_data:
    
    clearCSV(csv_file)
    
    def dataWriter(self, write):
        writer = csv.writer(open('same_test.csv', 'ab'), delimiter = ',')
        writer.writerow(write)
                         
    def OnData(self, Security, cookie, Fields, Data, Status):
        for i in range(0, bars):
            date = Data[i][-2]
            write = [str(date)]
            for sym in list:
                close = Data[i][-1]
                print 'Date: ' + sym + str(date) + "  Close: " + str(close)
                write.append(str(close))
            self.dataWriter(write)
    
    def OnStatus(self, Status, SubStatus, StatusDescription):
        print 'OnStatus'
        
class TestAsync:
    def __init__(self):
        clsid = '{F2303261-4969-11D1-B305-00805F815CBF}'
        progid = 'Bloomberg.Data.1'

        print 'connecting to BBCom........'
        print 'getting historical data.....'        
        blp = DispatchWithEvents(clsid, get_historical_data)
        blp.GetHistoricalData('AAPL US Equity', 1, entity, start, end, Results = Empty) 
        blp.AutoRelease = False
        blp.Flush()

        end_time = time() + 1
        
        while 1:
            PumpWaitingMessages()
            if end_time < time():
                print 'timed out'
                break
  
if __name__ == "__main__":
    ta = TestAsync()

