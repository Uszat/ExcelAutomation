import openpyxl
from openpyxl import Workbook
import os

#defines
OFFSET_TO_START_FROM_ONE = 1 #starts from 0 but I want to start from 1
OFFSET_TO_CALC_NO_PPL = 1 #people in sheet actually start from 2nd pos

#define shortcuts for the answers
tak = 'Tak, chcę zostać z kimś połączony/a'
nie = 'Nie - dołączam z własnym Fitness Buddy'
wmale = 'Mężczyzną'
wfemale = 'Kobietą'
male = 'Mężczyzna'
female = 'Kobieta'
unknown = 'Inna'
both = 'Dowolnie'
online = 'Online'
station = 'Stacjonarnie'
preferDate = 'Preferowanego terminu ćwiczeń w tygodniu'
preferSport = 'Preferowanej dyscypliny sportu'
rano = 'Rano (6-11)'
wdzien = 'W ciągu dnia (11-16)'
wieczor = 'Wieczorem (16-22)'
ranoiWdzien = 'Rano (6-11), W ciągu dnia (11-16)'
ranoiWieczor = 'Rano (6-11), Wieczorem (16-22)'
wdzieniWieczor = 'W ciągu dnia (11-16), Wieczorem (16-22)'
ranoiWdzieniWieczor = 'Rano (6-11), W ciągu dnia (11-16), Wieczorem (16-22)'
monday      = 0
tuesday     = 1
wednesday   = 2
thursday    = 3
friday      = 4
sat         = 5
sun         = 6

#sport dictionary
sport = {
    'yoga':'Yoga',
    'run':'Bieganie',
    'gym':'Siłownia',
    'fitness':'Zajęcia grupowe fitness (zumba',
    'bike':'Jazda na rowerze',
    'skates':'Jazda na rolkach',
    'homeExc':'Ćwiczenia w domu (np z filmikami na YT - Chodawkowska',
    'football':'Piłka nożna',
    'voleyball':'Siatkówka',
    'swim':'Pływanie',
    'dance':'Taniec',
    'tenis':'Tenis',
    'martial':'Sporty lub sztuki walki',
    'basketball':'Koszykówka'
    }

#specify the files location (or path)
cwd = os.path.dirname(os.path.abspath(__file__))
file = cwd + '\\fitness_buddy.xlsx'

#force shut the file fitness_buddy_appended
    #TODO ADD IT

#create an empty list to append values later on
values = []

#index of pairs with their buddy
pairIndex = 0

listOfGroups = []
hasPairGroup = []

#create a class of people
class Person(object):
    def __init__(self, number):
        self.number = number
        self.freqz = ''
        self.onlineStationSet = ''
        self.dateSportPreference = ''
        self.days = [[monday],[tuesday],[wednesday],[thursday],[friday],[sat],[sun]]
        self.genderBucket = 'notSpecified'

    def __eq__(self, other):
        if (self.looksForBuddy == other.looksForBuddy and 
            self.genderBucket == other.genderBucket and
            self.freqz == other.freqz and
            self.onlineStationSet == other.onlineStationSet
            ):
            return True
        else:
            return False

    #print person data
    def showData(self):
        print("number \t\t",        self.number + OFFSET_TO_START_FROM_ONE)
        print("name \t\t",          self.name)
        print("looksForBuddy \t",   self.looksForBuddy)
        print("nameToPair \t",      self.nameToPair)
        print("genderPair \t",      self.genderPair)
        print("onlineStation \t",   self.onlineStation)
        print("townFrom \t",        self.townFrom)
        print("dateSport \t",       self.dateSport)
        print("monday \t\t",        self.day[monday])
        print("tuesday \t",         self.day[tuesday])
        print("wednesday \t",       self.day[wednesday])
        print("thursday \t",        self.day[thursday])
        print("friday \t\t",        self.day[friday])
        print("sat \t\t",           self.day[sat])
        print("sun \t\t",           self.day[sun])
        print("discipline \t",      self.discipline)
        print("gender \t\t",        self.gender)
        print("freq \t\t",          self.freq)
        print("freqz \t\t",         self.freqz)
        print("genderBucket \t",    self.genderBucket)
        print(" ")

#open file and worksheet
workbook = openpyxl.load_workbook(file)
worksheet = workbook.active
noOfEntries = worksheet.max_row - OFFSET_TO_CALC_NO_PPL

#create new workbook for assigned pairs
wbAssigned = Workbook()
wsAssigned =  wbAssigned.active

#create list of objects People
people = []
for i in range(noOfEntries):
    people.append(Person(i))    

#fill People's attributes with values
def initObject():
    for index, person in enumerate(people):
        i = index + 2 #index starts from zero so make i from 2 as the real rows start from 2nd
        person.name =           worksheet['B' + str(i)].value
        person.looksForBuddy =  worksheet['D' + str(i)].value
        person.nameToPair =     worksheet['E' + str(i)].value
        person.genderPair =     worksheet['F' + str(i)].value
        person.onlineStation =  worksheet['G' + str(i)].value
        person.townFrom =       worksheet['H' + str(i)].value
        person.dateSport =      worksheet['I' + str(i)].value
        person.days[monday] =   worksheet['J' + str(i)].value
        person.days[tuesday] =  worksheet['K' + str(i)].value
        person.days[wednesday] =worksheet['L' + str(i)].value
        person.days[thursday] = worksheet['M' + str(i)].value
        person.days[friday] =   worksheet['N' + str(i)].value
        person.days[sat] =      worksheet['O' + str(i)].value
        person.days[sun] =      worksheet['P' + str(i)].value
        person.discipline =     worksheet['Q' + str(i)].value
        person.gender =         worksheet['R' + str(i)].value
        person.freq =           worksheet['S' + str(i)].value
        
initObject()

#setting all the variables for people
def setFrequency(person):
    if person.freq == 1 or person.freq == 2:
        person.freqz = 'low'
    elif person.freq == 3 or person.freq == 4:
        person.freqz = 'med'
    else:
        person.freqz = 'high'

def divideBySex(person):
    if person.gender == female and person.genderPair == wfemale:
        person.genderBucket = 'fwf'
    elif (person.gender == female and person.genderPair == wmale) or (person.gender == male and person.genderPair == wfemale):
        person.genderBucket = 'fwm'
    elif person.gender == male and person.genderPair == wmale:
        person.genderBucket = 'mwm'
    elif (person.gender == female and person.genderPair == both) or (person.gender == male and person.genderPair == both) or (person.gender == unknown and person.genderPair == both):
        person.genderBucket = 'both' #(f,m,u with both)
    elif (person.gender == unknown and person.genderPair == wfemale):
        person.genderBucket = 'uwf'
    elif (person.gender == unknown and person.genderPair == wmale):
        person.genderBucket = 'uwm'

def setOnlinePreference(person):
    if(person.onlineStation == online or person.onlineStation == both):
        person.onlineStationSet = online
    elif(person.onlineStation == station):
        person.onlineStationSet = station

times = [rano, wdzien, wieczor, ranoiWdzien, ranoiWieczor, ranoiWdzieniWieczor, wdzieniWieczor]

def setDateSportPreference(person):
    preferTime = {rano: 0, wdzien: 0, wieczor: 0, ranoiWdzien: 0, ranoiWieczor: 0, ranoiWdzieniWieczor: 0, wdzieniWieczor: 0}
    if(person.dateSport == preferDate):
        for day in range(7): #bo 7 dni
            for time in times:
                if(person.days[day] == time):
                    preferTime[time]+=1
        person.dateSportPreference = max(preferTime, key = preferTime.get) 
        print(person.dateSportPreference)
        #TODO nieskonczone, trzeba zakwalifikowac te slabe wyniki do czegos konkretnego, albo też przechodzic przez wszytkie ze wszytkimi i sprawdzac czy istnieje jakis match (taki matrix of matches)

    elif(person.dateSport == preferSport):
        person.dateSportPreference = person.discipline.split(sep=",")

#inserting people in excel cells
def putPersonCellShort(person):
    global pairIndex
    wsAssigned['A' + str(pairIndex)] = person.name 
    pairIndex+=1

def putPeopleInCell():
    global pairIndex         
    for id in range(len(listOfGroups)):
        for idd in range(len(listOfGroups[id])):
            pairIndex+=1
            wsAssigned['A' + str(pairIndex)] = listOfGroups[id][idd].number
            wsAssigned['B' + str(pairIndex)] = listOfGroups[id][idd].name
            wsAssigned['C' + str(pairIndex)] = listOfGroups[id][idd].looksForBuddy 
            wsAssigned['D' + str(pairIndex)] = listOfGroups[id][idd].nameToPair 
            wsAssigned['E' + str(pairIndex)] = listOfGroups[id][idd].genderPair 
            wsAssigned['F' + str(pairIndex)] = listOfGroups[id][idd].onlineStation 
            wsAssigned['G' + str(pairIndex)] = listOfGroups[id][idd].townFrom 
            wsAssigned['H' + str(pairIndex)] = listOfGroups[id][idd].dateSport 
            wsAssigned['I' + str(pairIndex)] = listOfGroups[id][idd].gender 
            wsAssigned['J' + str(pairIndex)] = listOfGroups[id][idd].freq  
        pairIndex+=1

def putHasBuddyPeopleInCell():
    global pairIndex         
    for id in range(len(hasPairGroup)):
        for idd in range(len(hasPairGroup[id])):
            pairIndex+=1
            wsAssigned['A' + str(pairIndex)] = hasPairGroup[id][idd].number
            wsAssigned['B' + str(pairIndex)] = hasPairGroup[id][idd].name
            wsAssigned['C' + str(pairIndex)] = hasPairGroup[id][idd].looksForBuddy 
            wsAssigned['D' + str(pairIndex)] = hasPairGroup[id][idd].nameToPair 
            wsAssigned['E' + str(pairIndex)] = hasPairGroup[id][idd].genderPair 
            wsAssigned['F' + str(pairIndex)] = hasPairGroup[id][idd].onlineStation 
            wsAssigned['G' + str(pairIndex)] = hasPairGroup[id][idd].townFrom 
            wsAssigned['H' + str(pairIndex)] = hasPairGroup[id][idd].dateSport 
            wsAssigned['I' + str(pairIndex)] = hasPairGroup[id][idd].gender 
            wsAssigned['J' + str(pairIndex)] = hasPairGroup[id][idd].freq  
        pairIndex+=1

def putPersonCellBelow(person):
    global pairIndex
    pairIndex+=1
    if(pairIndex % 3 == 0):
        pairIndex+=1     
    wsAssigned['A' + str(pairIndex)] = person.name 
    pairIndex+=1

#adding people to specific groups
def addPersonHasPair(person):
    global hasPairGroup
    pairListWasExtended = False
    for grpIdx, groupElements in enumerate(hasPairGroup):
        if(person.nameToPair == groupElements[0].name):
            hasPairGroup[grpIdx].extend([person])
            pairListWasExtended = True
            break
    if(pairListWasExtended == False):
        hasPairGroup.append([person])

def appendPeopleToList(person):
    global listOfGroups
    listWasExtended = False
    for grpIdx, allElementsOfGroup in enumerate(listOfGroups): #go through all elements of listofgroups 
        if(person == allElementsOfGroup[0]):
            #TODO Go through all sports from person and check them against all sports of personInAGroup - break on first match. Should be improved to get best match for current person
                #to get rid of single unpaired people but shit's too difficult ://
            #for item in person.dateSport
            #if person.dateSportPreference ==  
                #
            listOfGroups[grpIdx].extend([person]) #extend the current dimension of listofgroups (indicated by INDEX) by an element of current valueList element
            listWasExtended = True
            break
        elif(person.looksForBuddy == nie):
            addPersonHasPair(person)
            listWasExtended = True
            break
    if(listWasExtended == False):
        listOfGroups.append([person])

#main assigning loop                                     
def assignPeople():

    for person in people:
        #set person's frequency of excercises
        setFrequency(person)

        #rozdziel ludzi do kontenerow wzgledem plci i preferencji plci
        divideBySex(person)

        #set dowolnie preference to online
        setOnlinePreference(person)

        #set day preference or sport in one variable
        setDateSportPreference(person)

        #przydziel ludzi juz razem 
        appendPeopleToList(person)

#show all People's data
def showAllPeople():
    for person in people:
        person.showData()

assignPeople()
#showAllPeople()

putPeopleInCell()
putHasBuddyPeopleInCell()

#save new worksheet with added pairs
wbAssigned.save(cwd + '\\fitness_buddy_assigned.xlsx')


# for id in range(len(hasPairGroup)):
#     for idd in range(len(hasPairGroup[id])):
#         print("ids: ", id, idd)
#         hasPairGroup[id][idd].showData()

# for idx, lst in enumerate(hasPairGroup):
#     print("index", idx, "value: ", lst)
