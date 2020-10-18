import pandas as pd
import requests
from bs4 import BeautifulSoup
import re

def remove_brackets(line):
    regex = re.compile(".*?\((.*?)\)")
    result = re.findall(regex, line)
    return line.replace(result[0],'').replace('(','').replace(')','')

''' -------------------------------------- Harry Potter 4 scraping -------------------------------------- '''

url = 'http://nldslab.soe.ucsc.edu/charactercreator/film_corpus/film_20100519/all_imsdb_05_19_10/Harry-Potter-and-the-Goblet-of-Fire.html'
response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")
bTags=soup.findAll('b')
Characters=[]
Sentence=[]
for tag in bTags:
    tagText=tag.text.strip()
    if 'EXT.' not in tagText and 'INT.' not in tagText and tagText!='':
        sentenceText = str(tag.next_sibling).strip().replace('<b>', '').replace('</b>', '').partition("\n\n")[0].replace('\n',' ')
        if '(' in sentenceText:
            sentenceText=remove_brackets(sentenceText)
        Characters.append(tagText.capitalize())
        Sentence.append(sentenceText)

HarryPotter4={'Character':Characters, 'Sentence': Sentence}
df4 = pd.DataFrame(HarryPotter4, columns = ['Character', 'Sentence'])
#df4.to_excel("Harry_Potter_4_Script.xlsx")


''' -------------------------------------- Harry Potter 6 scraping -------------------------------------- '''

url = 'http://nldslab.soe.ucsc.edu/charactercreator/film_corpus/film_20100519/all_imsdb_05_19_10/Harry-Potter-and-the-Half-Blood-Prince.html'
response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")
bTags=soup.findAll('b')
Characters=[]
Sentence=[]

for tag in bTags:
    tagText=tag.text.strip()
    if 'EXT.' not in tagText\
            and 'INT.' not in tagText\
            and tagText!=''\
            and ':' not in tagText \
            and ')' not in tagText \
            and '.' not in tagText \
            and '-' not in tagText:
        sentenceText = str(tag.next_sibling).strip().replace('<b>', '').replace('</b>', '').partition("\n\n")[0].replace('\n',' ')
        if '(' in sentenceText:
            sentenceText=remove_brackets(sentenceText)
        Characters.append(tagText.capitalize())
        Sentence.append(sentenceText)

HarryPotter6={'Character':Characters, 'Sentence': Sentence}
df6 = pd.DataFrame(HarryPotter6, columns = ['Character', 'Sentence'])
#df6.to_excel("Harry_Potter_6_Script.xlsx")