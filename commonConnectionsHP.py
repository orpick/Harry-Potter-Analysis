import pandas as pd
from sqlalchemy import create_engine
import matplotlib.pyplot as plt
from operator import itemgetter
from collections import Counter
import unicodedata
import string
import pandasql as ps
from sqlalchemy.sql import text

engine = create_engine('sqlite://', echo=False)

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


''' --------------------------- creating a list containing lower-case characters' names --------------------------- '''
Characters=engine.execute("SELECT First_Name FROM Characters")
CharactersNames=pd.DataFrame(Characters, columns=['Name'])
CharactersLowerNames=CharactersNames.Name.str.lower()


''' ------------------------------------------ lower-casing a given text ------------------------------------------ '''
def normalize_caseless(text):
    return unicodedata.normalize("NFKD", text.casefold())


''' ----------- returning a word-frequency dictionary: (word : number of appearences in the input list) ----------- '''
def most_common(instances):
    """Returns a list of (instance, count) sorted in total order and then from most to least common"""
    return sorted(sorted(Counter(instances).items(), key=itemgetter(0)), key=itemgetter(1), reverse=True)


''' ----- creating a words list from the input characters lines and sending it to word-frequency calculation ------ '''
def charactersWordsDictionary(LinesBycharacterName):
    characterName_words = []
    for i in range(0, len(LinesBycharacterName)):
        line = LinesBycharacterName.loc[i].str.split(' ')[0]
        normalized_line = [normalize_caseless((line[index]).translate(str.maketrans('', '', string.punctuation))) for
                           index in range(0, len(line))]
        characterName_words.extend(normalized_line)
    return most_common(characterName_words)


''' ---------------------------- returning a character-references-frequency dictionary ---------------------------- '''
def common_connections(charactersDictionary):
    connection_dict={}
    for word in charactersDictionary:
        for name in CharactersLowerNames.iteritems():
            if word[0] == name[1]:
                connection_dict[word[0]]=word[1]
    return connection_dict


''' --- returning the 3 most refered friends (that are not Harry, Hermione or Ron) in movies 1,2,3,4,6 (refered by characterName) --- '''
def CharactersFriends(movieFriends):
    charactersFriends = {}
    for movie in movieFriends:
        for friend in movie.keys():
            if friend in charactersFriends.keys():
                charactersFriends[friend] += movie[friend]
            else:
                charactersFriends[friend] = movie[friend]

    charactersFriends = {k: v for k, v in sorted(charactersFriends.items(), key=lambda item: item[1], reverse=True)}
    charactersFriends = {'name': charactersFriends.keys(), 'number_references': charactersFriends.values()}
    charactersFriendsTable = pd.DataFrame.from_dict(charactersFriends)
    TopThreeFriends = "SELECT name, number_references\
                       FROM charactersFriendsTable \
                       WHERE name NOT IN ('harry', 'hermione', 'ron')\
                       LIMIT 3"
    TopThreeFriendsTable=ps.sqldf(TopThreeFriends)
    return TopThreeFriendsTable


''' ----------- creating the refered-characters dictionary according to Harry Potter 1,2,3,4,6 scripts ----------- '''
def charactersWordsByMovies(characterName):
    movies=['HarryPotter1Script', 'HarryPotter2Script', 'HarryPotter3Script','HarryPotter4Script','HarryPotter6Script']
    movieFriends=[]
    conn = engine.connect()
    for movie in movies:
        linesFetch = text('SELECT Line FROM ' + movie + ' WHERE Character= :characterName')
        SentencesByCharacter=conn.execute(linesFetch, characterName=characterName).fetchall()
        LinesByCharacter= pd.DataFrame(SentencesByCharacter, columns=['line'])
        movieFriends.append(common_connections(charactersWordsDictionary(LinesByCharacter)))
    charactersFriends = CharactersFriends(movieFriends)
    return charactersFriends


''' ---------------------------- counting a characters number of words in given script ----------------------------- '''
def countCharactersWords(tableName, characterName):
    countWords=0
    conn = engine.connect()
    linesFetch=text('SELECT Line FROM '+tableName+' WHERE Character= :characterName')
    SentencesByCharacter=conn.execute(linesFetch, characterName=characterName).fetchall()
    LinesByCharacter = pd.DataFrame(SentencesByCharacter, columns=['line'])
    Characters_words = charactersWordsDictionary(LinesByCharacter)
    for word in Characters_words:
        countWords += word[1]
    return countWords


''' ---- creating a counter array containing the number of words said by a given character in movies 1,2,3,4,6 ---- '''
def charactersWordsCounterArray(characterName):
    CountCharactersWords = []
    movies=['HarryPotter1Script', 'HarryPotter2Script', 'HarryPotter3Script','HarryPotter4Script','HarryPotter6Script']
    for movie in movies:
        CountCharactersWords.append(countCharactersWords(movie, characterName))
    return CountCharactersWords


''' ----------------------- returning the number of Voldemort refernces in a specific script ----------------------- '''
def VoldemortAppearences(tableName, movieNum):
    conn = engine.connect()
    if movieNum==1 or movieNum==3 or movieNum==4:
            Vappearences=text("SELECT COUNT(Line) AS num_V_references FROM "+tableName+" WHERE Line LIKE '%Voldemort%' or Line LIKE '% Dark Lord%'")
    elif movieNum==2:
        Vappearences = text("SELECT COUNT(Line) AS num_V_references FROM "+tableName+" WHERE Line LIKE '%Voldemort%' or Line LIKE '% Tom%' or Line LIKE '% Dark Lord%' or Line LIKE '% Riddle%'")
    else:
        Vappearences = text("SELECT COUNT(Line) AS num_V_references FROM "+tableName+" WHERE Line LIKE '%Voldemort%' or Line LIKE '% Tom %' or Line LIKE '% Dark Lord%'")
    VReferences=conn.execute(Vappearences , tableName=tableName).fetchall()
    VoldemortReferences= pd.DataFrame(VReferences, columns=['num_references'])
    return VoldemortReferences.num_references[0]


''' ------ returning an array containing the number of Voldemort references in Harry Potter 1,2,3,4,6 scripts ------ '''
def voldemortReferences():
    movies=['HarryPotter1Script', 'HarryPotter2Script', 'HarryPotter3Script','HarryPotter4Script','HarryPotter6Script']
    VoldemortArray = []
    for movie in range(len(movies)):
        VoldemortArray.append(VoldemortAppearences(movies[movie], movie+1))
    return VoldemortArray




''' --------------------------------------------------------------------------------------------------------------- '''
''' --------------------------------------------------- OUTPUTS --------------------------------------------------- '''
''' --------------------------------------------------------------------------------------------------------------- '''

''' --------- printing the top 3 refered-characters that Harry, Hermione and Ron refered to in movies 1,2,3,4,6 (excluding themselves) --------- '''
trioArray=["Harry", "Hermione", "Ron"]
for name in trioArray:
    print()
    print(name)
    charactersFriends=charactersWordsByMovies(name)
    print(charactersFriends)


''' --------------------------------------------- Voldemort References --------------------------------------------- '''

''' ------------------- plotting a graph that shows Voldemort's referenecs by movies (1,2,3,4,6) ------------------- '''
movies=[1,2,3,4,6]
VoldemortArray=voldemortReferences()
plt.plot(movies, VoldemortArray)
plt.title("Voldemort References By Movie")
plt.xlabel('Movie')
plt.ylabel('Number of Voldemort References')
plt.show()


''' --- plotting the number of words distribution throughout movies 1,2,3,4,6 of each chosen character on the same plot --- '''

countHermionesWords=charactersWordsCounterArray('Hermione')
countRonsWords=charactersWordsCounterArray('Ron')
countHarrysWords=charactersWordsCounterArray('Harry')
countDumbledoresWords=charactersWordsCounterArray('Dumbledore')
countHagridsWords=charactersWordsCounterArray('Hagrid')

plt.plot(movies, countHarrysWords, c='b', label='Harry')
plt.plot(movies, countRonsWords, c='r', label='Ron')
plt.plot(movies, countHermionesWords, c='g', label='Hermione')
plt.plot(movies, countDumbledoresWords, c='c', label='Dumbledore')
plt.plot(movies, countHagridsWords, c='m', label='Hagrid')

plt.legend(loc='upper left')
plt.title("Number of Words in Each Movie (Excluding Harry Potter 5)")
plt.xlabel('Movie')
plt.ylabel('Number of words')
plt.show()