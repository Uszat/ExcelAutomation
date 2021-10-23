#pip3 install openpyxl
from typing import ItemsView
import openpyxl

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
stationary = 'Stacjonarnie'
preferDays = 'Preferowanego terminu ćwiczeń w tygodniu'
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
file = 'D:/Downloads/fitness_buddy.xlsx'

#create an empty list to append values later on
values = []

#create a class of people
class Person(object):
    def __init__(self, number):
        self.number = number
        self.rareFreq = False
        self.medFreq = False
        self.highFreq = False
        self.genderBucket = 'notSpecified'

    #print person data
    def showData(self):
        print("number \t\t",        self.number + 1)
        print("name \t\t",          self.name)
        print("looksForBuddy \t",   self.looksForBuddy)
        print("nameToPair \t",      self.nameToPair)
        print("genderPair \t",      self.genderPair)
        print("onlineStation \t",   self.onlineStation)
        print("townFrom \t",        self.townFrom)
        print("baseOnDateSport ",   self.baseOnDateSport)
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
        print("freq \t\t",          self.rareFreq)
        print("freq \t\t",          self.medFreq)
        print("freq \t\t",          self.highFreq)
        print("genderBucket \t",    self.genderBucket)
        print(" ")

#open file and worksheet
workbook = openpyxl.load_workbook(file)
worksheet = workbook.active
noOfEntries = worksheet.max_row - 1

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
        person.baseOnDateSport =worksheet['I' + str(i)].value
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

def assignPeople():

    for person in people:
        #set person's frequency of excercises
        if person.freq == 1 or person.freq == 2:
            person.rareFreq = True
        elif person.freq == 3 or person.freq == 4:
            person.medFreq = True
        else:
            person.highFreq = True  

        #rozdziel ludzi do kontenerow wzgledem plci i preferencji plci
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

        #przydziel ludzi juz razem 
        if (person.genderBucket == 'fwf' and person.rareFreq and person.onlineStation == online and person.baseOnDateSport == preferDays):
            print(person.name)



# for person in people:
#     if(person.looksForBuddy == nie and person.nameToPair is not None):
#         print("")


#show all People's data
def showAllPeople():
    for person in people:
        person.showData()

assignPeople()


"""
Trzeba przejsc przez ludzi i powkładać ich do kontenerów
    - has friend
    - hasnt got friend
    
    if female w female and freq == 1 or freq == 2 and online
       1 put to f with f and freq 1-2
    if female w female and freq == 3 or freq == 4
       1 put to f with f and freq 3-4
    if female w female and freq == 5 or freq == 6 or freq == 7
       1 put to f with f and freq 5-7

    else if female with m or male and wants with f and freq == 1 or freq == 2 
       2 put to f with m and freq 1-2
    else if female with m or male and wants with f and freq == 3 or freq == 4 
       2 put to f with m and freq 3-4
    else if female with m or male and wants with f and freq == 5 or freq == 6 or freq == 7 
       2 put to f with m and freq 3-4
       
#tak jak wyzej czestotliwosc dla kolejnych 4 przypadkow 

    else if male and wants with m
       3 put to m with m
    else if f wants with b or m with b or u with b
       4 put to both
    else if u with f
       5 put to f with both
    else if u with m
       6 put to m with both 


    how many times a week 

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