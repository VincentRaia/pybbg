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

for i in list:
    symbol = i
    class get_historical_data:
        def dataWriter(self, date, close):
            writer = csv.writer(open(str(i)+'_historical_test.csv', 'ab'), delimiter = ',')
            data_to_enter = [str(date), str(close)]
            writer.writerow(data_to_enter)

        def merger():
            f1 = csv.reader(open('aapl_historical_test.csv', 'rb'))
            f2 = csv.reader(open('ibm_historical_test.csv', 'rb'))

            mydict = {}
            for row in f1:
                mydict[row[0]] = row[1:]

            for row in f2:
                mydict[row[0]] = mydict[row[0]].extend(row[1:])

            fout = csv.write(open('merged.csv','w'))
            for k,v in mydict:
                fout.write([k]+v)
                             
        def OnData(self, Security, cookie, Fields, Data, Status):
            for i in range(0, bars):
                date = Data[i][-2]
                close = Data[i][-1]
                print 'Date: ' + str(date) + "  Close: " + str(close)
                self.dataWriter(date, close)
        
        def OnStatus(self, Status, SubStatus, StatusDescription):
            print 'OnStatus'
            
    class TestAsync:
        def __init__(self):
            clsid = '{F2303261-4969-11D1-B305-00805F815CBF}'
            progid = 'Bloomberg.Data.1'

            print 'connectting to BBCom........'
            print 'getting historical data.....'        
            blp = DispatchWithEvents(clsid, get_historical_data)
            blp.GetHistoricalData(symbol, 1, entity, start, end, Results = Empty) 
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

