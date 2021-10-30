listOfGroups = [[1,1,1], [3, 3], [4,4], [5,5], [11], [13,13,13]]

# valueList = [[1],[1],[3],[4],[4],[6],[7],[8],[9],[10]] #zbędne, ja po prostu będę robił liste od 1 do 10 wiec wystarczy value in 10
# valueList = [[1],[2],3,4,5,6,7,8,9,10]
valueList = [[1],[2],[69],[69],[3], [55]]
#moj person to zwykla liczba, jesli liczba sie zgadza to ma takie same warunki??? chyba tak

useless = 0
#czyli valueList to moja lista persons (people)
# a listofGroups to jak tutaj, pierwotnie pusta, lista grup ktora sie tworzy poprzez appendowanie do niej ludzi lub extendowanie
for index, value in enumerate(valueList): #go through values of valueList (saving index of the value in INDEX)
    listWasExtended = False
    for grpIdx, allElementsOfGroup in enumerate(listOfGroups): #go through all elements of listofgroups 
        if(value[0] == allElementsOfGroup[0]): #if first element of value of valueList equals first element of current dimension of listofgroups then
            print("current value z listy", value[0]) 
            print("element: ", allElementsOfGroup[0])
            print("grupa przed dodaniem", listOfGroups)
            listOfGroups[grpIdx].extend(value) #extend the current dimension of listofgroups (indicated by INDEX) by an element of current valueList element
            print("grupa po dodaniu    ", listOfGroups)
            listWasExtended = True
            break
        else:
            useless += 1
            print("Useless", useless)
            #with break useless goes from 40 to 29
    if(listWasExtended == False):
        listOfGroups.append(value)
        print("Appended", value, "from index: ", index)
        print(listOfGroups)
        

            # break #zeby nie dodał jedne osoby do dwoch grup albo zbednie nie przechodzil przez grupy
            
print(listOfGroups)


#now make it for objects