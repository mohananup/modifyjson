# to read an excel and input the value in JSON
# and change the value of excel in round robin method


import collections
from collections import OrderedDict
import pandas as pd
from pandas import ExcelWriter
from openpyxl import Workbook
import json 

# Give the name of the file 
path = "devices.xlsx"

#read excel
rl= pd.read_excel(r'devices.xlsx')

print("Column \n", rl.columns)

# Set the JSON file path here 
PATH_TO_JSON = 'AndroidSuiteJson' 
 

#read existing json to memory.  
with open(PATH_TO_JSON,'r') as jsonfile:
    json_content = json.load(jsonfile, object_pairs_hook=OrderedDict) # this is now in memory! can be use it outside 'open'

#to get the list of devices
devicelist = rl['Devicelist']
print("Devices are below \n",devicelist)

#no of devices in the list
noofrows=len(rl.index)
print("no of rows are \n",noofrows)


for i in range(0, noofrows):
    
    json_content['configurations'][i] ['adapters'] [0] ['properties'] [7] ['value'] = devicelist[i]
    #print (json_content['configurations'][i] ['adapters'] [0] ['properties'] [7] ['value'])
 

with open(PATH_TO_JSON,'w') as jsonfile:
   json.dump(json_content, jsonfile, indent=6)
 

#rl.shape
noofrows=len(rl.index)
print("no of rows are \n",noofrows)

newdevicelist = []

#round robin logic
for i in range(0,noofrows-1):
    newdevicelist.append(devicelist[i+1])
    
newdevicelist.append(devicelist[0])
print("NEW device list after roundrobin \n",newdevicelist)

#converting list to dataframe
listdf = pd.DataFrame({'Devicelist':newdevicelist})

#writing back to excel file
writer1 = ExcelWriter('devices.xlsx')
listdf.to_excel(writer1,'Sheet1',index=False)
writer1.save()