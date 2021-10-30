#pip3 install openpyxl
from typing import ItemsView
import openpyxl
from openpyxl import Workbook
import os
import numpy as np

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

#sport dictionary
sport = {
    'yoga':'Yoga, rozciąganie',
    'run':'Bieganie, jogging',
    'gym':'Siłownia',
    'fitness':'Zajęcia grupowe fitness (zumba, aerobic, crossfit)',
    'bike':'Jazda na rowerze',
    'skates':'Jazda na rolkach, wrotkach, hulajnodze, deskorolce',
    'homeExc':'Ćwiczenia w domu (np z filmikami na YT - Chodawkowska, Agata Zająć, Blogilates)',
    'football':'Piłka nożna',
    'voleyball':'Siatkówka',
    'swim':'Pływanie',
    'dance':'Taniec',
    'tenis':'Tenis, Squash, badminton',
    'martial':'Sporty lub sztuki walki',
    'basketball':'Koszykówka'
    }

#specify the files location (or path)
cwd = os.path.dirname(os.path.abspath(__file__))
file = cwd + '\\fitness_buddy.xlsx'

#create an empty list to append values later on
values = []

#index of pairs with their buddy
pairIndex = 1

listOfGroups = []

#create a class of people
class Person(object):
    def __init__(self, number):
        self.number = number
        self.freqz = ''
        # self.lowFreq = False
        # self.medFreq = False
        # self.highFreq = False
        self.genderBucket = 'notSpecified'

    def __eq__(self, other):
        if (self.looksForBuddy == other.looksForBuddy and 
            self.genderBucket == other.genderBucket and
            self.freqz == other.freqz and
            self.onlineStation == other.onlineStation and
            self.dateSport == other.dateSport):
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
        print("monday \t\t",        self.monday)
        print("tuesday \t",         self.tuesday)
        print("wednesday \t",       self.wednesday)
        print("thursday \t",        self.thursday)
        print("friday \t\t",        self.friday)
        print("sat \t\t",           self.sat)
        print("sun \t\t",           self.sun)
        print("discipline \t",      self.discipline)
        print("gender \t\t",        self.gender)
        print("freq \t\t",          self.freq)
        print("freqz \t\t",         self.freqz)
        # print("freq \t\t",          self.lowFreq)
        # print("freq \t\t",          self.medFreq)
        # print("freq \t\t",          self.highFreq)
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
        person.monday =         worksheet['J' + str(i)].value
        person.tuesday =        worksheet['K' + str(i)].value
        person.wednesday =      worksheet['L' + str(i)].value
        person.thursday =       worksheet['M' + str(i)].value
        person.friday =         worksheet['N' + str(i)].value
        person.sat =            worksheet['O' + str(i)].value
        person.sun =            worksheet['P' + str(i)].value
        person.discipline =     worksheet['Q' + str(i)].value
        person.gender =         worksheet['R' + str(i)].value
        person.freq =           worksheet['S' + str(i)].value
        
initObject()

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

def putPersonCellShort(person):
    global pairIndex
    wsAssigned['A' + str(pairIndex)] = person.name 
    pairIndex+=1

def putPersonCell(person):
    global pairIndex     
    wsAssigned['A' + str(pairIndex)] = person.name
    wsAssigned['B' + str(pairIndex)] = person.looksForBuddy 
    wsAssigned['C' + str(pairIndex)] = person.nameToPair 
    wsAssigned['D' + str(pairIndex)] = person.genderPair 
    wsAssigned['E' + str(pairIndex)] = person.onlineStation 
    wsAssigned['F' + str(pairIndex)] = person.townFrom 
    wsAssigned['G' + str(pairIndex)] = person.dateSport 
    wsAssigned['H' + str(pairIndex)] = person.gender 
    wsAssigned['I' + str(pairIndex)] = person.freq   
    # if(pairIndex % 3 == 0):
    #     pairIndex+=2
    # else:
    #     pairIndex+=1    

def putPersonCellBelow(person):
    global pairIndex
    pairIndex+=1
    if(pairIndex % 3 == 0):
        pairIndex+=1     
    wsAssigned['A' + str(pairIndex)] = person.name 
    pairIndex+=1

def pairFriends(person):
    global pairIndex
    if(person.looksForBuddy == nie and person.nameToPair is not None):
        putPersonCell(person)
        pairIndex+=1

def pairRest(person):
    global pairIndex
    genderList = ['fwf', 'fwm', 'mwm', 'both', 'uwf', 'uwm']
    freqList = ['low', 'med', 'high']
    onlineStationList = [online, station, both]
    dateSportList = [preferDate, preferSport]
    #tego na razie nie uzywam groupList = np.zeros((6,3,3,2))

    if(person.looksForBuddy == tak):
        for gender in genderList:
            if(gender == person.genderBucket):
                for freq in freqList:
                    if(freq == person.freqz):
                        for onlineStation in onlineStationList:
                            if(onlineStation == person.onlineStation):
                                for dateSport in dateSportList:
                                    if(dateSport == person.dateSport):
                                        putPersonCell(person)
                                        for p in listOfGroups:
                                            if(person == p):
                                                listOfGroups[0].extend(person)
                                            else:
                                                listOfGroups.append(person)
                                            #add the person under this stack list
                                       # else:
                                       #     add it as a new stack new array 
                                        pairIndex+=1
#make list = [group1], [group2]... i jesli person == person from any group of the list then add person do tej grupy, else add do nowej grupy


def pairRestIDIOT(person):
    if(person.genderBucket == 'fwf'):
        if(person.freqz == 'low'):
            if(person.onlineStation == online or person.onlineStation == both):
                if(person.dateSport == preferDate):
                    putPersonCellBelow(person)
                elif(person.dateSport == preferSport):
                    putPersonCellBelow(person)
            elif(person.onlineStation == station):
                #if there is sb from your town ? proceed : goto online
                if(person.dateSport == preferDate):
                    putPersonCellBelow(person)
                elif(person.dateSport == preferSport):
                    putPersonCellBelow(person)
        elif(person.freqz == 'med'):
            if(person.onlineStation == online or person.onlineStation == both):
                if(person.dateSport == preferDate):
                    putPersonCellBelow(person)
                elif(person.dateSport == preferSport):
                    putPersonCellBelow(person)
            elif(person.onlineStation == station):
                #if there is sb from your town ? proceed : goto online
                if(person.dateSport == preferDate):
                    putPersonCellBelow(person)
                elif(person.dateSport == preferSport):
                    putPersonCellBelow(person)
        elif(person.freqz == 'high'):
            if(person.onlineStation == online or person.onlineStation == both):
                if(person.dateSport == preferDate):
                    putPersonCellBelow(person)
                elif(person.dateSport == preferSport):
                    putPersonCellBelow(person)
            elif(person.onlineStation == station):
                #if there is sb from your town ? proceed : goto online
                if(person.dateSport == preferDate):
                    putPersonCellBelow(person)
                elif(person.dateSport == preferSport):
                    putPersonCellBelow(person)

def assignPeople():

    for person in people:
        #set person's frequency of excercises
        setFrequency(person)

        #rozdziel ludzi do kontenerow wzgledem plci i preferencji plci
        divideBySex(person)

        #przydziel ludzi co maja kumpla
        pairFriends(person)
        
        #przydziel ludzi juz razem 
        pairRest(person)
        #pairRestIDIOT(person)

#show all People's data
def showAllPeople():
    for person in people:
        person.showData()

assignPeople()
#showAllPeople()

wbAssigned.save(cwd + '\\fitness_buddy_assigned.xlsx')









"""

Trzeba zrobić grupy - 48 grup chyba
moze objekt grupa albo cooooooooooś i automatycznie przypisuje osobe do danej grupy 
w zależności od tego w jakie fory loopy itp poszedl

    kontenery:
    - plec
        - f w f
        - m w f
        - m w m
        - f,m,u w b
        - f w b
        - m w b
    - freq
        - 1,2
        - 3,4
        - 5,6,7
    - online or stationary
        -online
        -stationary
        -both
    - Town (tu trzeba zaimplementowac poprawianie znaków)
        - Lodz
        - WWa
        - Pobliskie???????
    - Time vs Discipline
        - Time
        - Discipline
"""