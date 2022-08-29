from asyncio.windows_events import NULL
import json
from openpyxl import load_workbook
import requests
import pandas as pd
import html
# import cgi, os
# import cgitb; cgitb.enable()
# form = cgi.FieldStorage()
# fileitem = form['filename']
 
# # check if the file has been uploaded
# if fileitem. :
#     # strip the leading path from the file name
#     fn = os.path.basename(fileitem.filename)
     
#    # open read and write the file into the server

# Removing last space for every word to match up the activity
def removeLastSpace(word):
    if word[len(word) -1] == ' ':
        word = word.split(" ")[0]
    return word

# Altering the value of each column based on the column name.
def alterValue(word, key = ''):
    if word == "nan":
        return ''
    elif key == 'indoor':
        if word and word == 'Both':
            word = 'IndoorOutdoorBoth'
    elif key == 'land':
        if word and word == 'Both':
            word = 'LandWaterBoth'
        elif word == 'Land activity':
            return 'Land'
        elif word == 'Water activity':
            word = 'Water'
    elif key == 'individual':
        if word and word == 'Both':
            word = 'individualGroupBoth'
    elif key == 'child':
        if word and word == 'Both':
            word = 'ChildrenAdultBoth'
    elif key == 'access':
        print("fkjd")
        word = 3
    return word

# Latitude and Longitude aoi request params key and URL.
params = {
  "access_key": 'e78a0264da0e76b7dbc3953b4a3d7cf2',
  "query": ''
}
url = "http://api.positionstack.com/v1/forward"

file = pd.ExcelFile('python_task.xlsx')
data = pd.read_excel(file, 'Upated Jefferson_Shelby_YK AB')
act = pd.read_excel(file, 'Instructions')
master = pd.DataFrame(act)
masterActivities = pd.DataFrame(act, columns= ['Master activity type list'])
items = list(masterActivities.values.flatten())
fd = pd.DataFrame(data)
fd = fd.applymap(lambda x: html.unescape(str(x)).encode().decode())
df = pd.DataFrame(data, columns= ['Address', 'Activities', 'childrenAdultBoth', 'individualGroupBoth', ''])
df = df.applymap(lambda x: html.unescape(str(x)).encode().decode())
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
                if act not in items:
                    if act not in newItems:
                        newItems.insert(index, act)
                        index = index + 1
    except Exception as e:
        print("Err", e)
#     res = requests.get(url, params)
#     if res.status_code == 200:
#         data = json.loads(res._content.decode("utf-8"))
#         if data['data'] and data['data'][0]:
#             coord = str(data['data'][0]['latitude']) + ", " + str(data['data'][0]['longitude'])
#             coordinates.insert(index, coord)
#         else: 
#             coordinates.insert(index, item[0])
#     else:
#         coordinates.insert(index, item[0])

# coordinates = { 'Address': coordinates}
# cdn = pd.DataFrame(coordinates)
# fd['Address'] = cdn['Address']

fd[['indoorOutdoorBoth']] = fd[['indoorOutdoorBoth']].applymap(lambda x: str(alterValue(x)))
fd[['childrenAdultBoth']] = fd[['childrenAdultBoth']].applymap(lambda x: alterValue(x, 'child'))
fd[['individualGroupBoth']] = fd[['individualGroupBoth']].applymap(lambda x: alterValue(x, 'individual'))
fd[['LandWaterBoth']] = fd[['LandWaterBoth']].applymap(lambda x: alterValue(x, 'land'))
fd[['Accessibility Score']] = fd[['Accessibility Score']].applymap(lambda x: alterValue(x, 'access'))
fd=fd.dropna()
newItems = { 'Master activity type list': newItems }
dfi = pd.DataFrame(newItems)
finalAct = pd.concat([master, dfi])
finalAct.to_excel("abc.xlsx", index=False)
fd.to_csv('fileMap.csv', index=False)