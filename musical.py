#Version One

#connect to google doc

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json


scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

credentials = ServiceAccountCredentials.from_json_keyfile_name("api-project-325394400367-0c233ad31779.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials) # authenticate the JSON key with gspread
sheet = file.open("MusiCal: Willy Wonka Kids") #open sheet


#for each row in the schedule:

sheet = sheet.worksheet("Schedule")
for i in range(1,sheet.row_count):
    r = sheet.row_values(i)
    if (i < 4):
        continue
    print (r)
    (date,st,et,s1,s2,s3,s4,s5,typ1,typ2,x) = r
    print(date,st,et,s1,typ1)



    #find date and time
    #create list of scenes
#create calendar entry for each rehersal 



#Version Two

#parse scenes/songs to create a list of scenes2characters 
#parse cast assignments to create a list of characters2student
#parse character groups to create a list of groups2characters
#for each scene in list of scenes/songs:
    #from scenes/songs get the list of characters required for that scene
    #


