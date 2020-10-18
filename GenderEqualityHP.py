import pandas as pd
from sqlalchemy import create_engine
import matplotlib.pyplot as plt
from operator import itemgetter
from collections import Counter
import unicodedata
from sqlalchemy.sql import text
import numpy as np

engine = create_engine('sqlite://', echo=False)
movies=[1,2,3,4,6]

''' ---------------------------------------------- returning a table ---------------------------------------------- '''
def createTable(filePath, tableName):
    df = pd.read_excel(filePath)
    df.to_sql(tableName, engine)
    return df


''' ----------------------------------- creating the tables needed for the code ----------------------------------- '''
createTable('C:\\Users\\OR\\Downloads\\HarryPotterCharacters.xlsx', 'Characters')
createTable('C:\\Users\\OR\\Downloads\\Harry_Potter_1_Script.xlsx', 'HarryPotter1Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\Harry_Potter_2_Script.xlsx', 'HarryPotter2Script')
createTable('C:\\Users\\OR\\Downloads\\Harry_Potter_3_Script.xlsx', 'HarryPotter3Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\\Harry_Potter_4_Script.xlsx', 'HarryPotter4Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\\Harry_Potter_6_Script.xlsx', 'HarryPotter6Script')


''' -------------------------------------------- General Harry Potter  -------------------------------------------- '''
CharactersByGender=engine.execute("SELECT Gender, COUNT(Gender) AS num_characters FROM Characters\
                                    WHERE Gender IS NOT NULL GROUP BY Gender")
GenderAndCharacters=pd.DataFrame(CharactersByGender, columns=['Gender', 'num_characters'])
female_characters=GenderAndCharacters.loc[0, 'num_characters']
male_characters=GenderAndCharacters.loc[1,'num_characters']

#for every 2 male characters, only one female character
print("female/male ratio in Harry Potter characters: ", female_characters/male_characters)
print()


''' ------------------------------------------------------ R ------------------------------------------------------ '''

''' ------------ Creation of excel files containing each movie's characters and their number of lines ------------- '''
''' ------------------------------ also returns the number of female and male lines ------------------------------- '''
def genderLinesSpecificMovie(tableName, charactersColumnName, movieNum, excelName):
    HPSentencesByGender = engine.execute("SELECT Gender, COUNT(Line) AS num_lines FROM\
                                          Characters JOIN " + tableName + " ON " + charactersColumnName + "=" + tableName + ".Character" \
                                         + " GROUP BY Gender")
    HPLinesByGender = pd.DataFrame(HPSentencesByGender, columns=['Gender', 'num_lines'])
    HPFemaleLines = HPLinesByGender.loc[0, 'num_lines']
    HPMaleLines = HPLinesByGender.loc[1, 'num_lines']

    print("female/male ratio in Harry Potter", movieNum, "lines: ", HPFemaleLines / HPMaleLines)
    print()

    AllGendersLines = engine.execute("SELECT Character, COUNT(Line) AS num_lines, Gender FROM\
                                          Characters JOIN " + tableName + " ON " + charactersColumnName + "=" + tableName + ".Character\
                                          GROUP BY " + tableName + ".Character ORDER BY Gender")
    AllGendersLinesTable = pd.DataFrame(AllGendersLines, columns=['Name', 'num_lines', 'Gender'])
    # AllGendersLinesTable.to_excel(excelName)
    return (HPFemaleLines, HPMaleLines)


''' --------------------- fills ratio list and female and male lines counter lists with values --------------------- '''
def genderLinesAllMovies(scriptsNames, columnNames, movieNumbers, excelNames, ratioList, FLines, MLines):
    for i in range(len(scriptsNames)):
        (femaleLines, maleLines)=genderLinesSpecificMovie(scriptsNames[i], columnNames[i], movieNumbers[i], excelNames[i])
        ratioList.append(femaleLines/maleLines)
        FLines.append(femaleLines)
        MLines.append(maleLines)


''' ------------------- calculating the number of female and male characters in a specific movie ------------------- '''
def charactersByGender(tableName, charactersColumnName, movieNum):

    conn = engine.connect()
    linesFetch=text('SELECT Gender, COUNT(Gender) as num_characters FROM\
                                (SELECT DISTINCT Character, Gender \
                                FROM '+tableName+' JOIN Characters\
                                 ON '+tableName+'.Character='+charactersColumnName+') GROUP BY Gender')
    maleFemaleCharacters=conn.execute(linesFetch, tableName=tableName, charactersColumnName=charactersColumnName).fetchall()
    print("Female/Male Characters in Harry Potter", movieNum)
    HPGenderAndCharacters=pd.DataFrame(maleFemaleCharacters, columns=['Gender', 'num_characters'])
    print(HPGenderAndCharacters)
    return HPGenderAndCharacters.num_characters


''' ------------------------------------- chosen characters words distribution ------------------------------------- '''

''' ------------------------------------------ lower-casing a given text ------------------------------------------- '''
def normalize_caseless(text):
    return unicodedata.normalize("NFKD", text.casefold())


''' ------------ returning a word-frequency dictionary: (word : number of appearences in the input list) ----------- '''
def most_common(instances):
    """Returns a list of (instance, count) sorted in total order and then from most to least common"""
    return sorted(sorted(Counter(instances).items(), key=itemgetter(0)), key=itemgetter(1), reverse=True)


''' ----- filling two lists (list per gender) with the number of characters of each gender in movies 1,2,3,4,6 ----- '''
def numCharByGenderLists(scriptsNames, columnNames, movieNumbers, female_characters_number, male_characters_number):
    for i in range(len(scriptsNames)):
        num_characters = charactersByGender(scriptsNames[i], columnNames[i], movieNumbers[i])
        female_characters_number.append(num_characters[0])
        male_characters_number.append(num_characters[1])




''' --------------------------------------------------------------------------------------------------------------- '''
''' --------------------------------------------------- OUTPUTS --------------------------------------------------- '''
''' --------------------------------------------------------------------------------------------------------------- '''


''' -------------------- plotting the number of characters throughout movies 1,2,3,4,6 by gender --------------------'''
scriptsNames=['HarryPotter1Script', 'HarryPotter2Script', 'HarryPotter3Script','HarryPotter4Script','HarryPotter6Script']
columnNames=['Characters.Fname', 'Characters.Fname2', 'Characters.Fname3', 'Characters.Fname4', 'Characters.Fname6']
excelNames=['LinesByGenderHP1.xlsx', 'LinesByGenderHP2.xlsx', 'LinesByGenderHP3.xlsx', 'LinesByGenderHP4.xlsx', 'LinesByGenderHP6.xlsx']
movieNumbers=[1,2,3,4,6]
female_characters_number = []
male_characters_number = []

''' filling male_characters_number and female_characters_number lists with values '''
numCharByGenderLists(scriptsNames, columnNames, movieNumbers, female_characters_number, male_characters_number)

plt.plot(movies, male_characters_number, c='b', label='male characters')
plt.plot(movies, female_characters_number, c='r', label='female characters')
plt.legend(loc='upper right')
plt.title("Number of Characters by Gender in Each Movie (Excluding Harry Potter 5)")
plt.xlabel('Movie')
plt.ylabel('Number of characters')
plt.yticks(np.arange(0, max(male_characters_number)+1, 2))
plt.show()


''' ---------------------------- plotting the number of lines by each gender per movie ----------------------------- '''
ratioList=[] #this was used for later R calculations
FLines=[]
MLines=[]

''' filling ratioList, FLines and MLines lists with values '''
genderLinesAllMovies(scriptsNames, columnNames, movieNumbers, excelNames, ratioList, FLines, MLines)
plt.plot(movies, MLines, c='b', label='male lines')
plt.plot(movies, FLines, c='r', label='female lines')
plt.legend(loc='upper right')
plt.title("Lines by Gender in Each Movie (Excluding Harry Potter 5)")
plt.xlabel('Movie')
plt.ylabel('Number of lines')
plt.show()