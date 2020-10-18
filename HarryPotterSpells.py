import pandas as pd
from sqlalchemy import create_engine
from operator import itemgetter
from collections import Counter
import unicodedata
import string
from sqlalchemy.sql import text

engine = create_engine('sqlite://', echo=False)

''' ---------------------------------------------- returning a table ---------------------------------------------- '''
def createTable(filePath, tableName):
    df = pd.read_excel(filePath)
    df.to_sql(tableName, engine)
    return df

''' ----------------------------------- creating the tables needed for the code ----------------------------------- '''

createTable('C:\\Users\\OR\\Downloads\\HarryPotterCharacters.xlsx', 'Characters')
createTable('C:\\Users\\OR\\Downloads\\HarryPotterSpells.xlsx', 'Spells')
createTable('C:\\Users\\OR\\Downloads\\Harry_Potter_1_Script.xlsx', 'HarryPotter1Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\Harry_Potter_2_Script.xlsx', 'HarryPotter2Script')
createTable('C:\\Users\\OR\\Downloads\\Harry_Potter_3_Script.xlsx', 'HarryPotter3Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\\Harry_Potter_4_Script.xlsx', 'HarryPotter4Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\\Harry_Potter_6_Script.xlsx', 'HarryPotter6Script')


''' ----------------------------------------- lower-casing all the spells ----------------------------------------- '''
AllSpells=engine.execute("SELECT Spell FROM Spells WHERE Spell IS NOT 'Unknown'")
Spells=pd.DataFrame(AllSpells, columns=['Spell'])
lowerSpells=Spells.Spell.str.lower()


''' -------------------- creating a table consists of all spells containing more than one word -------------------- '''
twoWordSpells=engine.execute("SELECT Spell FROM Spells WHERE Spell LIKE '% %'")
twoWordSpells=pd.DataFrame(twoWordSpells, columns=['line'])


''' ------------------------------------------ lower-casing a given text ------------------------------------------ '''
def normalize_caseless(text):
    return unicodedata.normalize("NFKD", text.casefold())


''' ----------- returning a word-frequency dictionary: (word : number of appearences in the input list) ----------- '''
def most_common(instances):
    """Returns a list of (instance, count) sorted in total order and then from most to least common"""
    return sorted(sorted(Counter(instances).items(), key=itemgetter(0)), key=itemgetter(1), reverse=True)


''' ----------------- adding to a spell dictionary all the spells that contain more than one word ----------------- '''
def twoWordSpellsfunc(twoWordSpells, Lines, spell_dic):
    for line in Lines.line:
        for spell in twoWordSpells.line:
            if spell in line:
                if spell in spell_dic.keys():
                    spell_dic[spell] += 1
                else:
                    spell_dic[spell] = 1

    return spell_dic


''' --------------- normallizing a characters' lines in order to create a word-frequency dictionary --------------- '''
def words_distribution(Lines):
    words = []
    for i in range(0, len(Lines)):
        line = Lines.loc[i].str.split(' ')[0]
        normalized_line = [normalize_caseless((line[index]).translate(str.maketrans('', '', string.punctuation))) for
                           index in range(0, len(line))]
        words.extend(normalized_line)

    return most_common(words)


''' ---------------------------------------- returning a spells dictionary ---------------------------------------- '''
def spellsDictionary(tableName, characterName=None):
    if characterName==None:
        sentences = engine.execute("SELECT Line FROM " + tableName + " WHERE Line IS NOT NULL")
    else:
        conn = engine.connect()
        linesFetch=text('SELECT Line FROM '+tableName+' WHERE Character= :characterName AND Line IS NOT NULL')
        sentences=conn.execute(linesFetch, characterName=characterName).fetchall()
    Lines = pd.DataFrame(sentences, columns=['line'])
    words = words_distribution(Lines)
    return creating_spell_dict(words, Lines, lowerSpells, twoWordSpells)


''' ----------------------------------------- creating a spells dictionary ---------------------------------------- '''
def creating_spell_dict(words, Lines, lowerSpells, twoWordSpells):
    spell_dic = {}

    for dist in words:
        if dist[0] in lowerSpells.values:
            spell_dic[dist[0]] = dist[1]
    spell_dic=twoWordSpellsfunc(twoWordSpells, Lines, spell_dic)
    return sorted(sorted(spell_dic.items(), key=itemgetter(0)), key=itemgetter(1), reverse=True)


''' -------------------------- for every movie in moviesList, finding its most used spell ------------------------- '''
def mostUsedSpells(moviesList):
    mostUsedSpellsAllMovies=[]
    timesUsed=[]
    for movie in moviesList:
        #mostUsedSpells list will contain the most used spells in the movie
        mostUsedSpells=[]
        #movieSpells contains a sorted (descending order) spells dictionary
        movieSpells=spellsDictionary(movie)
        #extracting the most used spell in the movie, and the amount of times it was used
        mostUsedSpell=movieSpells[0]
        #adding the first most used spell to the mostUsedSpells list
        mostUsedSpells.append(mostUsedSpell[0])
        #adding the number of times that the first most used spell was used to timesUsed list
        timesUsed.append(mostUsedSpell[1])
        #if there are several most-used-spells, adding them to the list
        for i in range(1,len(movieSpells)):
            if movieSpells[i][1]==mostUsedSpell[1]:
                mostUsedSpells.append(movieSpells[i][0])
            else:
                break
        #adding the most used spells in the movie to the overall list of most used spells in all the movies
        mostUsedSpellsAllMovies.append(mostUsedSpells)
    return (mostUsedSpellsAllMovies, timesUsed)


''' ----------------------- for every spell in spells list, finding its most common caster ------------------------ '''
def mostCaster(tableName, spells):
    casters=[]
    for spell in spells:
        Findcaster = engine.execute("SELECT Character FROM \
                                (SELECT Character, MAX(numCastings) FROM \
                                (SELECT Character, COUNT(Line) AS numCastings FROM "+tableName+" \
                                 WHERE Line LIKE '%"+spell+"%' GROUP BY Character))")
        caster= pd.DataFrame(Findcaster, columns=['casterName'])
        casters.append(caster.casterName[0])
    return casters


''' -------- returning a dataFrame that holds for every movie in moviesList its most used spell and caster -------- '''
def moviesAndFavSpells(moviesList, moviesNames):
    casters=[]
    i=0
    favSpells=mostUsedSpells(moviesList)
    mostUsedSpellsAllMovies=favSpells[0]
    timesUsed=favSpells[1]

    #for every most used spell, finding the character that casted it the most
    for spells in mostUsedSpellsAllMovies:
        casters.append(mostCaster(moviesList[i], spells))
        i+=1

    #creating a dataFrame that will hold all the information collected about each movies most used spells
    data = {'movie': moviesNames,
            'most used spell(s)': mostUsedSpellsAllMovies,
            'number of times used': timesUsed,
            'common caster': casters}

    moviesAndSpells= pd.DataFrame(data, columns=['movie', 'most used spell(s)', 'number of times used', 'common caster'])
    return moviesAndSpells


''' ---- creating a sorted dictionary of spells used by a specific character throughout all movies 1,2,3,4,6 ----- '''
def charactersSpellsDictAllMovies(name, movies):
    finalCharacterSpellDict={}
    for movie in movies:
        movieCharacterSpellDict=spellsDictionary(movie, name)
        for spell in movieCharacterSpellDict:
            if spell[0] in finalCharacterSpellDict.keys():
                finalCharacterSpellDict[spell[0]]+=spell[1]
            else:
                finalCharacterSpellDict[spell[0]]=spell[1]
    return sorted(sorted(finalCharacterSpellDict.items(), key=itemgetter(0)), key=itemgetter(1), reverse=True)


''' ----------------------------------- creating a partial dictionary of spells ----------------------------------- '''
def mostUsedSpellsByCharacter(spellsDict):
    mostUsedSpells={}
    numCastings=spellsDict[0][1]
    spellsCounter=1
    mostUsedSpells[spellsDict[0][0]]=numCastings
    for i in range(1, len(spellsDict)):
        if spellsDict[i][1]!=numCastings:
            spellsCounter+=1
            if spellsCounter==4:
                break
        numCastings = spellsDict[i][1]
        mostUsedSpells[spellsDict[i][0]] = numCastings
    return mostUsedSpells




''' --------------------------------------------------------------------------------------------------------------- '''
''' --------------------------------------------------- OUTPUTS --------------------------------------------------- '''
''' --------------------------------------------------------------------------------------------------------------- '''


''' --------------- printing the spells dictionary of chosen characters throughout movies 1,2,3,4,6 --------------- '''
names=['Harry', 'Hermione', 'Ron', 'Snape']
movies = ['HarryPotter1Script', 'HarryPotter2Script', 'HarryPotter3Script', 'HarryPotter4Script', 'HarryPotter6Script']
movieNum=1
for movie in movies:
    print("Harry Potter ", movieNum, " Spells:")
    print()
    for name in names:
        print(name,"'s spell dictionary: ", spellsDictionary(movie, name))
    movieNum+=1
    print()
    print()


''' --------- printing a dataFrame that holds for every movie in moviesList its most used spell and caster --------- '''
''' ------- Note: "Harry Potter 4" is not in the list because no spells were written in this movie's script -------- '''
moviesList = ['HarryPotter1Script', 'HarryPotter2Script', 'HarryPotter3Script', 'HarryPotter6Script']
moviesNames=['Harry Potter 1', 'Harry Potter 2', 'Harry Potter 3', 'Harry Potter 6']
df=moviesAndFavSpells(moviesList, moviesNames)
#print(df)
#df.to_excel("CommonSpellsAndCastersByMovie.xlsx", index=False)


''' ------------- for each character in names list, printing a partial list of their most used spells ------------- '''
names=['Harry', 'Hermione', 'Ron', 'Snape']
movies = ['HarryPotter1Script', 'HarryPotter2Script', 'HarryPotter3Script', 'HarryPotter6Script']
for name in names:
    spellDict=charactersSpellsDictAllMovies(name, movies)
    print(name, ": ", mostUsedSpellsByCharacter(spellDict))