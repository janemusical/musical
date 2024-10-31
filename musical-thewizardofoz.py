# Theater rehearsal scheduling system that processes Google Sheets data to generate individual and full cast calendars
# 
# Required spreadsheet format:
# - Schedule worksheet: Date, Times, Scenes/Songs (6 columns), Rehearsal Types (6 columns), Location
# - Scenes worksheet: Scene name, Act number (Roman), Characters (9 columns)
# - Characters/Ensembles worksheet: Character name, Act 1/2 dialogue flags, Ensemble flags
# - Cast Assignments worksheet: Student name, Character assignments (10 columns)

#Version One

#uid needs to be unique  

# Required imports for:
# - Google Sheets API access and authentication
# - Calendar generation (.ics files)
# - Date/time handling with timezone support
# - String parsing and file operations
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime
from parse import *
from icalendar import Calendar, Event, vCalAddress, vText
from pathlib import Path
import os
import pytz
import uuid

# Core class representing a single rehearsal session
# Tracks who needs to be there, what's being rehearsed, when and where
# Handles formatting of rehearsal details for both display and calendar export
class Rehearsal:
    def __init__(self, startdt, enddt, what, who, where):
        self.startdt = startdt    # datetime with timezone - start time
        self.enddt = enddt        # datetime with timezone - end time 
        self.what = what          # list of (scene/song, rehearsal type) tuples
        self.who = who            # list of student names required for this rehearsal
        self.where = where        # location string

    # Creates HTML-formatted string of rehearsal details for display
    # Format: Date/Time in bold, followed by scenes/songs and participants
    def __str__(self):
        # html tags to bold the Date information
        r = "<b>" + ppDT(self.startdt, self.enddt) + "</b>"
        # start a new line for What is rehearsed
        r = r + "<br>" + "\nWhat: "

        # List each scene/song with its rehearsal type (e.g., "Vocal: Opening Number")
        for w, t in self.what:
            r = r + t + ": " + w + ", "

        # Remove the last 2 characters to get rid of ", " at the end
        r = r[:-2]

        # Add list of all required participants
        r = r + "<br>" + "\nWho: "

        for w in self.who:
            r = r + w + ", "

        # Remove the last 2 characters to get rid of ", " at the end
        # Add two lines between each entry
        r = r[:-2] + "<br><br>"

        # This line adds the location of the rehearsal to the print out
        #r = r + "\n" + self.where
        return r

# Class for managing contact information for cast/crew members
# Currently used for future email functionality
class Person:
    def __init__(self, name, email1, email2):
        self.name = name          # Full name of person
        self.email1 = email1      # Primary email (e.g., student email)
        self.email2 = email2      # Secondary email (e.g., parent email)

    def __str__(self):
        return str(self.name) + " - " + self.email1 + ": " + self.email2

# Parses date and time strings from spreadsheet into datetime objects
# Handles AM/PM conversion and sets Central timezone
# Date format: MM/DD/YYYY
# Time format: HH:MM AM/PM
def createTimes (d, t):
    print (d,t)
    # Parse date components
    (mm, dd, yyyy) = parse("{:d}/{:d}/{:d}", d)
    # Parse time components including AM/PM
    (hour, minute, daytime) = parse("{:d}:{:d} {}", t)
    #print (mm, dd, yyyy, hour, minute, daytime)
    # Convert to 24-hour format if PM
    if daytime == 'PM' and hour != 12:
        hour +=12
    
    # Create datetime with Central timezone
    return datetime(yyyy, mm, dd, hour, minute, tzinfo=pytz.timezone('America/Chicago'))   #TODO source timezone

# Formats date/time range into readable string for display
# Example output: "August 13, 2024 from 03:00 PM - 05:00 PM"
def ppDT(sdt,edt):
    # Get start date
    pretty_start_date_time = sdt.strftime("%B %d, %Y from %I:%M %p")  # "August 13, 2024 at 03:00 PM"
  
    # Format the date into a pretty string
    pretty_date = "" + pretty_start_date_time + " - " + edt.strftime("%I:%M %p")
    return pretty_date

# Creates an iCalendar event object for a specific rehearsal and student
# Includes formatted description of scenes/songs being rehearsed
def createEvent (r, student):
    event = Event()
    studentLengthNumber = len(student)
    # Extract first name for event title
    sss = student.split(" ")
    studentFirstName = sss[0]
    #studentFirstName = student[:(studentLengthNumber - 2)]
    
    # Set event title with show name and student
    event.add('summary', 'The Wizard of Oz - ' + studentFirstName) #TODO hardcoded name
    
    # Create HTML-formatted description of rehearsal content
    desc = ''
    for scene,stype in r.what:
        #desc += scene + ': ' + stype + '\n'  ## Old Way, now adding formatting and switched order Type: Scene/Song
        desc += stype + ': ' + '<b><i>' + scene + '</b></i>' + '\n'
    
    # Add all required event fields
    event.add('description', desc)
    event.add('dtstart', r.startdt)
    event.add('dtend', r.enddt)
    organizer = vCalAddress('MAILTO:janemargaretkramer@gmail.com') #TODO 
    organizer.params['name'] = vText('ACTT') #TODO
    event['organizer'] = organizer
    event['location'] = vText(r.where) 
    event['uid'] = 'ww' + str(uuid.uuid4()) #TODO make it a hash         
    event.add('priority', 5)
    
    return event

```python
# Processes Schedule worksheet to create list of Rehearsal objects
# Maps scenes to required characters and then to assigned students
# Handles both individual character assignments and full scene cast lists
def generateRehearsals (sheet, ScenestoChars, CharstoStudents):
    # Get all rows from Schedule worksheet including headers
    schedule = sheet.worksheet("Schedule").get_all_values()
    
    rehearsals = []

    #for each row in the schedule:
    i = 0
    for r in schedule:
        i += 1
        # Skip first 3 header rows
        if (i < 4):
            continue

        #find date and time
        # Parse 17 columns from schedule:
        # - date, start time, end time
        # - 6 scene/song columns
        # - 6 rehearsal type columns 
        # - location
        # - notes (x)
        (date,st,et,s1,s2,s3,s4,s5,s6,typ1,typ2,typ3,typ4,typ5,typ6,loc,x) = r

        # Convert string dates/times to datetime objects
        start_dt = createTimes (date, st)
        end_dt   = createTimes (date, et)

        #create list of scenes or characters
        # Get non-empty scenes from the 6 possible scene columns
        scenes_tmp = (s1,s2,s3,s4,s5,s6)
        scenes = []
        for s in scenes_tmp:
            if s != "":
              scenes.append(s)

        #create list of type of rehersal
        # Get non-empty rehearsal types matching the scenes
        rehersal_type_tmp = (typ1,typ2,typ3,typ4,typ5,typ6)
        types = []
        for t in rehersal_type_tmp:
            if t != "":
              types.append(t)
    
        #create rehearsal
        # for each scene
        # Build list of all required participants
        people = []
        
        for sceneorchar in scenes:
            # Check if this is a direct character assignment or a scene
            # is this a character or a scenes
            if sceneorchar in CharstoStudents:
                chars = [sceneorchar]  # Direct character assignment
            else:
                #   look up the characters required from scenes2chars
                chars = ScenestoChars[sceneorchar]  # Get all characters in scene

            #   for each character get the persons who play that role
            for char in chars:
                #   add that the list of people required for that rehearsal
                # Get all students assigned to this character
                students = CharstoStudents[char]
                for s in students:
                    #p = Person(s, 'jaydeekay@gmail.com', 'janemargaretkramer@gmail.com')
                    people.append(s)
    
        #dedupe the list
        print(len(people))
        #create a list that is both the scene and the type together to use in the rehersal text
        #eg Vocal: Act English
        # Zip scenes and their rehearsal types together
        combined_list = list(zip(scenes, types))
        print(combined_list)                                                                                          
        
        # Remove duplicate student names since they might appear multiple times
        # if they play multiple characters in the same scene
        people = list(dict.fromkeys(people))    
        print(len(people))

        # Create Rehearsal object with all collected information
        r = Rehearsal(start_dt, end_dt, combined_list, people, loc)
        #should we add types as another argument to rehersal or should it be embedded within scenes and also should we add location here too
        rehearsals.append(r)

    return rehearsals

# Processes Scenes worksheet to create two mappings:
# 1. ScenestoChars: scene name -> list of required characters
# 2. ActstoScenes: act number -> list of scenes in that act
def generateScenestoChars (sheet):
    # Get all rows from Scenes worksheet
    scenes = sheet.worksheet("Scenes").get_all_values()
    ScenestoChars = {}   # Dictionary mapping scene names to list of required characters
    ActstoScenes = {}    # Dictionary mapping act numbers to list of scenes
    ActstoScenes[1] = [] # Initialize empty lists for both acts
    ActstoScenes[2] = []

    #for each row in the scenes:
    i = 0
    for r in scenes:
        i += 1
        # Skip first 3 header rows
        if (i < 4):
            continue
        if not r:
            break

        # Parse row: scene name, act number, up to 9 character columns, notes
        (scene, act, c1, c2, c3, c4, c5, c6, c7, c8, c9, x) = r

        #create list of characters and ensembles
        # Get all non-empty character names from the 9 possible columns
        chars_tmp = (c1,c2,c3, c4, c5, c6, c7, c8, c9)
        chars = []
        for c in chars_tmp:
            if c != "":
              chars.append(c)
            
        # Map scene name to its list of required characters
        ScenestoChars[scene] = chars

        #if appropriate, add this to the ActstoScenes
        # If act number is specified, add scene to that act's list
        if act != '':
            ActstoScenes [int(act)].append(scene)
            
    return ScenestoChars, ActstoScenes

# Processes Characters/Ensembles worksheet to determine which characters
# have dialogue in each act
def generateActDialogueLists(sheet):
    # Get all rows from Characters/Ensembles worksheet
    CharactersAndEnsembles = sheet.worksheet("Characters / Ensembles").get_all_values()
    
    # Dictionary mapping act numbers to lists of characters with dialogue
    ActtoCharsWithDialogueInThatAct = {}   # Act -> list of chacters who have dialogue in that act
    ActtoCharsWithDialogueInThatAct[1] = []
    ActtoCharsWithDialogueInThatAct[2] = []

    #for each row in the charactes
    i = 0
    for r in CharactersAndEnsembles:
        i += 1
        if (i < 4):
            continue
        if not r:
            break
        #print (r)
        # Parse row: character name, Act 1 dialogue flag, Act 2 dialogue flag, ensemble flags
        (character, act1dialogue, act2dialogue, ensemblesinact, inprogram) = r

        # Add character to appropriate act lists if they have dialogue
        if act1dialogue == "1":
            ActtoCharsWithDialogueInThatAct[1].append(character)
        if act2dialogue == "1":
            ActtoCharsWithDialogueInThatAct[2].append(character)
        
    return ActtoCharsWithDialogueInThatAct
        
# Processes Cast Assignments worksheet to create:
# 1. CharstoStudents: character name -> list of assigned students
# 2. Students: complete list of all students in show                             
def generateCharstoStudents (sheet):
    # Get all rows from Cast Assignments worksheet
    castAssignments = sheet.worksheet("Cast Assignments").get_all_values()
    
    CharstoStudents = {}   # char -> list of students
    Students = []          # Complete list of student names

    #for each row in the cast assignments:
    i = 0
    for r in castAssignments:
        i += 1
        # Skip first 3 header rows
        if (i < 4):
            continue
        if not r:
            break
        #print (r)
        # Parse row: student name, up to 10 character assignments
        (student, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, x) = r
        Students.append(student)
        # Get all non-empty character assignments
        chars_tmp = (c1, c2, c3, c4, c5, c6, c7, c8, c9, c10)
        chars = []
        for c in chars_tmp:
            if c != "":
              chars.append(c)
        
    # for each character in that row
    #   check if we already have alist for that character, if so, get it and add the student
    #                                                      else, create a new list for that character and add the student
        # For each character assignment, add student to that character's list
        for c in chars:
            if c in CharstoStudents:
                CharstoStudents[c].append(student)
            else:
                CharstoStudents[c] = [student]

    return CharstoStudents, Students

##
## MAIN EXECUTION SECTION
##

# Google Sheets API access scopes needed for reading spreadsheet
scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

# Initialize Google Sheets connection
credentials = ServiceAccountCredentials.from_json_keyfile_name("api-project-325394400367-0c233ad31779.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials) # authenticate the JSON key with gspread
sheet = file.open("MusiCal: The Wizard of Oz - Viz") #open sheet

#get all the rehearsals from the schedule

# Build mapping of which characters have dialogue in each act
ActtoCharsWithDialogueInThatAct = generateActDialogueLists(sheet)
print(ActtoCharsWithDialogueInThatAct)

# Build mappings of scenes to characters and acts to scenes
ScenestoChars, ActstoScenes = generateScenestoChars(sheet)
print ('Hello')
print (ScenestoChars)
print (ActstoScenes)
print ("\n")

# Process special scene types that need additional handling:
# 1. "Run Act X" - combines all characters from all scenes in that act
# 2. "Act X Dialogue" - uses pregenerated dialogue character lists
 # within the ScenestoChars, there might be two types of special Scenes, that will need custom logic to create the
# proper list of Characters.
#
# Act XXX Dialogue
# Run Act XXX
#
# Where, amusingly, XXX will be a Roman Numeral Act number.

# Starting with Run Act XXX:
#  - get all scene in that act.
#  - create a list from ScenestoChars that merges all the Chars in that Act's scenes. 
for scene in ScenestoChars:
    # Handle "Run Act" special scenes
    if "Run Act" in scene:
        # Parse Roman numeral act number
        actrn = parse("Run Act {}", scene)
        if actrn[0] == "I":
            Act = 1
        else:
            Act = 2

        # Combine character lists from all scenes in this act
        tempChars = []
        for s in ActstoScenes[Act]:
            tempChars = tempChars + ScenestoChars[s]

        #now override the ScenestoChars for this particular scene
        ScenestoChars[scene] = tempChars

# section for act xxx dialogue scenes
#
    # Handle "Act X Dialogue" special scenes
    if "Dialogue" in scene and "All" not in scene:
        print (scene)
        # Parse Roman numeral act number
        actrn = parse("Act {} Dialogue", scene)
        if actrn[0] == "I":
            Act = 1
        else:
            Act = 2        
            
        # Use pre-generated list of characters with dialogue in this act
        # conveniently there is already a list of all the characters in a given act.
        ScenestoChars[scene] = ActtoCharsWithDialogueInThatAct[Act]

print ("\n \n")
print (ScenestoChars)
                                                                                                                    
# Build mapping of characters to assigned students
CharstoStudents,Students = generateCharstoStudents(sheet)
print (CharstoStudents)

# Generate all rehearsal objects
rehearsals = generateRehearsals(sheet, ScenestoChars, CharstoStudents)

# Print rehearsal details for verification
for r in rehearsals:
    print(r)
    print("\n")

# Initialize calendar storage
StudentCalendars = {}
    
# create a calendar for each actor
# Create empty calendar for each student
for student in Students:
    cal = Calendar()
    cal.add('prodid', '-//My calendar product//example.com//')
    cal.add('version', '2.0')
    StudentCalendars[student] = cal
    
#create calendar entry for each rehearsal 
# Create master calendar for full cast
fullCastCal = Calendar()
fullCastCal.add('prodid', '-//My calendar product//example.com//')
fullCastCal.add('version', '2.0')

StudentCalendars["Complete"] = fullCastCal  # "Complete" will show up in Calendar

# Add each rehearsal to appropriate student calendars
# and to master calendar
# for each rehearsal, add it to the actors' calendar + the Full Cast calendar.
for r in rehearsals:
    # Add to each participating student's calendar
    for student in r.who:
        StudentCalendars[student].add_component(createEvent(r,student))
    # Add the event to the main calendar containing all events
    StudentCalendars["Complete"].add_component(createEvent(r,"Complete"))

# Create output directory for calendar files
directory = Path.cwd() / 'MyCalendar'
try:
   directory.mkdir(parents=True, exist_ok=False)
except FileExistsError:
   print("Folder already exists")
else:
   print("Folder was created")

# Write individual student calendar files
for student in Students:
    cal = StudentCalendars[student]
    # Format student name as "FirstName L" for filename
    s = student.split(" ")
    stud = s[0] + " " + s[1][0]
    
    # Create .ics file
    filename = 'The Wizard of Oz - ' + stud + ".ics"
    f = open(os.path.join(directory, filename), 'wb')
    f.write(cal.to_ical())
    f.close()

#print out the full cast calendar
student = 'Complete'  #"Complete" will show up in Calendar
cal = StudentCalendars[student]
filename = 'The Wizard of Oz - Complete Schedule.ics'
f = open(os.path.join(directory, filename), 'wb')
f.write(cal.to_ical())
f.close()
