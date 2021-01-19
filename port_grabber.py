#Imports needed modules
import requests
import csv
import json
import pandas as pd
import os
from csv import writer
from csv import reader
import pyexcel as pe

#checks for the csv file and deletes it if it exists
if os.path.exists('ports.csv'):
    print("Checking for csv files. ")
    print("Checking for csv files.. ")
    print("Checking for csv files... ")
    print("Checking for csv files.... ")
    print("Files exists removing! ")
    os.remove('ports.csv')
    print("Files has been removed! ")
else:
    print("Checking for csv files. ")
    print("Checking for csv files.. ")
    print("Checking for csv files... ")
    print("Checking for csv files.... ")
    print("Files do not exist! Resuming Script! ")

#sets variables for the GET request
url = "https://api.meraki.com/api/v0/organizations/<ORG ID>/inventory"
payload = None
headers = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "X-Cisco-Meraki-API-Key": "<API KEY>"
}

#makes a GET request using the variable defined earlier
response = requests.request('GET', url, headers=headers, data = payload)

#put the GET request response into a variable to be used for looping
inventory = response.json()

#creates a list from the inventory variable from devices with names that start with MS
switches = [device for device in inventory if device['model'][:2] in ('MS*')]

#creates a list to store the names of all the switches
swNames = []

#creates a list to store the output from the for loop below
switch_networks = []

#loops through each object in the switches variable and adds the serial number to the switch_network list if its not there already
for switch in switches:
    if switch['serial'] not in switch_networks:
        switch_networks.append(switch['serial'])
        swNames.append(switch['name'])

#prints the total amount of switches found in the given network
print('Found a total of %d switches configured!' % (len(switches)))


#creates a dataframe to store the output from the second get requests

with open('ports.csv', 'w') as fp:
    w = csv.writer(fp)
    w.writerow(['port number','name','tags','enabled','poeEnabled','type','vlan','voiceVlan','allowedVlans','isolationEnabled','rstpEnabled','stpGuard','linkNegotiation','portScheduleId','udld','accessPolicyNumber','macWhitelist','stickyMacWhitelist','stickyMacWhitelistLimit','Switch Name'])
    fp.close()
    pass

names_dict = dict(zip(switch_networks, swNames))

#loops through each object appending the serial number to the api url and makes a get request for each object
for key, value in names_dict.items():
    df_marks=pd.DataFrame()
    surl = f"https://api.meraki.com/api/v0/devices/{key}/switchPorts"
    sresponse = requests.request('GET', surl, headers=headers, data = payload)
    #if the status code is not 200 for ok skip this item and move to the next in the for loop
    if sresponse.status_code != 200:
        print ('Cannot connect to' + value + 'skipping and moving on to next item')
        continue

    #prints information to the console so the user doesn't fall asleep :)
    print("Sending a GET Request to" ,surl)
    print("Returning data from" ,value)

    #appends the returned json data to a pandas data frame
    df_marks=df_marks.append(sresponse.json(), ignore_index=False)
    df_marks['Switch Name'] = value
    csv_data = df_marks.to_csv('ports.csv',mode='a',index=False, header=False)

#Does some CSV Modifications
delList=[8,9,10,11,12,13,14] 
sheet = pe.load('ports.csv')
del sheet.row[1]
sheet.delete_columns(delList)
sheet.save_as('ports.csv')
sheet.close()

direct = os.getcwd()

print('Script has completed successfully! CSV file is located at ' + direct + '\ports.csv' )




