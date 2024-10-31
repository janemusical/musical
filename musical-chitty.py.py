#Version One


#title is missing
#uid needs to be unique
#connect to google doc

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from   datetime import datetime
from parse import *
from icalendar import Calendar, Event, vCalAddress, vText
from pathlib import Path
import os
import pytz
import uuid


class Rehearsal:
    def __init__(self, startdt, enddt, what, who):
        self.startdt = startdt
        self.enddt = enddt
        self.what = what
        self.who = who

    def __str__(self):
        r = "" + str(self.startdt) + " - " + str(self.enddt) + "\nWhat:\t"

        for w in self.what:
            r = r + ": " + w + "\t"
        r = r + "\nWho:\t"
        
        for w in self.who:
            r = r + ": " + w + "\t"
            
        return  r

class Person:
    def __init__(self, name, email1, email2):
        self.name = name
        self.email1 = email1
        self.email2 = email2       

    def __str__(self):
        return str(self.name) + " - " + self.email1 + ": " + self.email2


def createTimes (d, t):
    print (d,t)
    (mm, dd, yyyy) = parse("{:d}/{:d}/{:d}", d)
    (hour, minute, daytime) = parse("{:d}:{:d} {}", t)
    #print (mm, dd, yyyy, hour, minute, daytime)
    if daytime == 'PM' and hour != 12:
        hour +=12
    
    return datetime(yyyy, mm, dd, hour, minute, tzinfo=pytz.timezone('America/Chicago'))   #TODO source timezone

def createEvent (r, student):

    event = Event()
    studentLengthNumber = len(student) 
    studentFirstName = student[:(studentLengthNumber - 2]
    event.add('summary', 'Chitty Chitty Bang Bang - ' + studentFirstName) #TODO hardcoded name
    desc = ''
    for scene in r.what:
        desc += scene + '\n'
    event.add('description', desc)
    event.add('dtstart', r.startdt)
    event.add('dtend', r.enddt)
    organizer = vCalAddress('MAILTO:janemargaretkramer@gmail.com') #TODO 
    organizer.params['name'] = vText('Mr. Strohmeyer') #TODO
    event['organizer'] = organizer
    event['location'] = vText('De Smet Jesuit High School') #TODO
    event['uid'] = 'ww' + str(uuid.uuid4()) #TODO make it a hash         
    event.add('priority', 5)
    
    return event

def generateRehearsals (sheet, ScenestoChars, CharstoStudents):

    schedule = sheet.worksheet("Schedule").get_all_values()
    
    rehearsals = []

    #for each row in the schedule:
    i = 0
    for r in schedule:
        i += 1
        if (i < 4):
            continue

        #find date and time
        (date,st,et,s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,typ1,typ2,typ3,typ4,typ5,typ6,typ7,typ8,typ9,typ10,typ11,loc,x) = r

        start_dt = createTimes (date, st)
        end_dt   = createTimes (date, et)

        #create list of scenes or characters
        scenes_tmp = (s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11)
        scenes = []
        for s in scenes_tmp:
            if s != "":
              scenes.append(s)

        #create list of type of rehersal
        rehersal_type_tmp = (typ1,typ2,typ3,typ4,typ5,typ6,typ7,typ8,typ9,typ10,typ11)
        types = []
        for t in rehersal_type_tmp:
            if t != "":
              types.append(t)
    
    
        #create rehearsal
        # for each scene

        people = []
        
        for sceneorchar in scenes:

            # is this a character or a scenes
            if sceneorchar in CharstoStudents:
                chars = [sceneorchar]
            else:
            #   look up the characters required from scenes2chars
                chars = ScenestoChars[sceneorchar]

            #   for each character get the persons who play that role
            for char in chars:
                
            #   add that the list of people required for that rehearsal
                students = CharstoStudents[char]
                for s in students:
                    #p = Person(s, 'jaydeekay@gmail.com', 'janemargaretkramer@gmail.com')
                    people.append(s)
    
        #dedupe the list
        print(len(people))
        
        people = list(dict.fromkeys(people))    
        print(len(people))
        r = Rehearsal(start_dt, end_dt, scenes, people)
        #should we add types as another argument to rehersal or should it be embedded within scenes and also should we add location here too
        rehearsals.append(r)

    return rehearsals


def generateScenestoChars (sheet):

    scenes = sheet.worksheet("Scenes").get_all_values()
    
    ScenestoChars = {}

    #for each row in the scenes:
    i = 0
    for r in scenes:
        i += 1
        if (i < 4):
            continue
        if not r:
            break

        (scene, act, sceneNumber, c1, c2, c3, c4, c5, c6, c7, x) = r

        #create list of characters and ensembles

        chars_tmp = (c1,c2,c3,c4,c5,c6,c7)
        chars = []
        for c in chars_tmp:
            if c != "":
              chars.append(c)
            
        ScenestoChars[scene] = chars

    return ScenestoChars


def generateCharstoStudents (sheet):

    castAssignments = sheet.worksheet("Cast Assignments").get_all_values()
    
    CharstoStudents = {}   # char -> list of students
    Students = []

    #for each row in the cast assignments:
    i = 0
    for r in castAssignments:
        i += 1
        if (i < 4):
            continue
        if not r:
            break
        #print (r)
        (student, c1, c2, c3, c4, c5,c6,c7,c8,c9, x) = r
        Students.append(student)
        chars_tmp = (c1, c2, c3, c4, c5,c6,c7,c8,c9)
        chars = []
        for c in chars_tmp:
            if c != "":
              chars.append(c)
        
    # for each character in that row
    #   check if we already have alist for that character, if so, get it and add the student
    #                                                      else, create a new list for that character and add the student

        for c in chars:
            if c in CharstoStudents:
                CharstoStudents[c].append(student)
            else:
                CharstoStudents[c] = [student]

    return CharstoStudents, Students




##
## MAIN PART
##


scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]



credentials = ServiceAccountCredentials.from_json_keyfile_name("api-project-325394400367-0c233ad31779.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials) # authenticate the JSON key with gspread
sheet = file.open("MusiCal: Chitty Chitty Bang Bang SSP") #open sheet

#get all the rehearsals from the schedule


ScenestoChars = generateScenestoChars(sheet)
print (ScenestoChars)
print ("\n")
CharstoStudents,Students = generateCharstoStudents(sheet)
print (CharstoStudents)

rehearsals = generateRehearsals(sheet, ScenestoChars, CharstoStudents)

for r in rehearsals:
    print(r)
    print("\n")

StudentCalendars = {}
    
# create a calendar for each actor
for student in Students:
    cal = Calendar()
    cal.add('prodid', '-//My calendar product//example.com//')
    cal.add('version', '2.0')
    StudentCalendars[student] = cal
    
    
#create calendar entry for each rehearsal 
fullCastCal = Calendar()
fullCastCal.add('prodid', '-//My calendar product//example.com//')
fullCastCal.add('version', '2.0')

StudentCalendars["Full Cast"] = fullCastCal


# for each rehearsal, add it to the actors' calendar + the Full Cast calendar.
for r in rehearsals:
    for student in r.who:
        StudentCalendars[student].add_component(createEvent(r,student))
    # Add the event to the calendar
    StudentCalendars["Full Cast"].add_component(createEvent(r,"Full Cast"))

# Write to disk
directory = Path.cwd() / 'MyCalendar'
try:
   directory.mkdir(parents=True, exist_ok=False)
except FileExistsError:
   print("Folder already exists")
else:
   print("Folder was created")


for student in Students:
    cal = StudentCalendars[student]
    
    filename = 'Willy Wonka - ' + student + ".ics"
    f = open(os.path.join(directory, filename), 'wb')
    f.write(cal.to_ical())
    f.close()

#print out the full cast calendar
student = 'Full Cast' 
cal = StudentCalendars[student]
filename = 'Chitty Chitty Bang Bang - Complete Schedule.ics'
f = open(os.path.join(directory, filename), 'wb')
f.write(cal.to_ical())
f.close()






#Version Two

#parse scenes/songs to create a list of ScenestoCharacters 
#parse cast assignments to create a list of characters2student
#parse character groups to create a list of groups2characters
#for each scene in list of scenes/songs:
    #from scenes/songs get the list of characters required for that scene
    #


