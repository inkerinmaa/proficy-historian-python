#Install pywin32 - pip install pywin32
#Install PrettyTable - pip install PrettyTable

import PyADO
from prettytable import PrettyTable 

conn = PyADO.connect(None,host='TESTCOLLECT',user='abcdef',password='ghijklm',provider='iHOLEDB.iHistorian.1')
curs = conn.cursor()

curs.execute("SELECT timestamp, value, quality, tagname FROM ihrawdata WHERE samplingmode=rawbytime AND timestamp>='01-Mar-2019 13:58' AND (tagname=CELL1.TBM_010_ProductCode)") 
result = curs.fetchall()
descr = curs.description

header = [i[0] for i in descr]
table = PrettyTable(header)
for row in result:
    table.add_row(row)
print(table)

curs.close()
conn.close()
