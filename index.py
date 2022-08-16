import json
import string
from openpyxl import load_workbook
import requests
import pandas as pd
import sys
params = {
  "access_key": 'e78a0264da0e76b7dbc3953b4a3d7cf2',
  "query": ''
}
url = "http://api.positionstack.com/v1/forward"
path = ''
if not sys.argv[1]:
    print("Please provide path to xlsx file while running index file")  
    sys.exit()
else:
    path = sys.argv[1]
file = pd.ExcelFile('D$_Mapping_Data.xlsx')
data = pd.read_excel(file, 'Upated Jefferson_Shelby_YK AB')
act = pd.read_excel(file, 'Instructions')
master = pd.DataFrame(act)
masterActivities = pd.DataFrame(act, columns= ['Master activity type list'])
fd = pd.DataFrame(data)
df = pd.DataFrame(data, columns= ['Address'])

coordinates = []
index = 0
print("dkjfj", df, masterActivities)
for item in df.values:
    params['query'] = item[0]
    res = requests.get(url, params)
    if res.status_code == 200:
        data = json.loads(res._content.decode("utf-8"))
        if data['data'] and data['data'][0]:
            coord = str(data['data'][0]['latitude']) + ", " + str(data['data'][0]['longitude'])
            coordinates.insert(index, coord)
        else: 
            coordinates.insert(index, item[0])
    else:
        coordinates.insert(index, item[0])
    index = index + 1
fd['coordinates'] = coordinates
fd.replace('N/A', '')
fd.to_excel("file.xlsx")