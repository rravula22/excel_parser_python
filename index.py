from asyncio.windows_events import NULL
import json
from pickle import TRUE
import string
from openpyxl import load_workbook
import requests
import pandas as pd
import sys
import html

def removeLastSpace(word):
    if word[len(word) -1] == ' ':
        word = word.split(" ")[0]
    return word
params = {
  "access_key": 'e78a0264da0e76b7dbc3953b4a3d7cf2',
  "query": ''
}
url = "http://api.positionstack.com/v1/forward"
path = ''

file = pd.ExcelFile('python_task.xlsx')
data = pd.read_excel(file, 'Upated Jefferson_Shelby_YK AB')
act = pd.read_excel(file, 'Instructions')
master = pd.DataFrame(act)
masterActivities = pd.DataFrame(act, columns= ['Master activity type list'])
items = list(masterActivities.values.flatten())
index = len(items)
fd = pd.DataFrame(data)
fd = fd.applymap(lambda x: html.unescape(str(x)).encode().decode())
df = pd.DataFrame(data, columns= ['Address', 'Activities'])
fd.replace('N/A', '')
coordinates = []
index = 0
newItems = []
for item in df.values:
    params['query'] = item[0]
    try:
        if isinstance(item[1], str): 
            acts = item[1].split(",")
            for act in acts:
                act = removeLastSpace(act)
                if [act] not in masterActivities.values:
                    if act not in newItems:
                        items.insert(index, act)
                        index = index + 1
    except Exception as e:
        print("Err", e)
    # res = requests.get(url, params)
    # if res.status_code == 200:
    #     data = json.loads(res._content.decode("utf-8"))
    #     if data['data'] and data['data'][0]:
    #         coord = str(data['data'][0]['latitude']) + ", " + str(data['data'][0]['longitude'])
    #         coordinates.insert(index, coord)
    #     else: 
    #         coordinates.insert(index, item[0])
    # else:
    #     coordinates.insert(index, item[0])
# print(items, len(items))

# fd['coordinates'] = coordinates
items = { 'Master activity type list': items }
dfi = pd.DataFrame(items)

master['Master activity type list'] = dfi['Master activity type list']

master.to_excel("abc.xlsx")
fd.to_csv('fileMap.csv')