import pandas as pd
from sqlalchemy import create_engine
import matplotlib.pyplot as plt
import networkx as nx
from networkx.algorithms import bipartite

engine = create_engine('sqlite://', echo=False)

''' ---------------------------------------------- returning a table ---------------------------------------------- '''
def createTable(filePath, tableName):
    df = pd.read_excel(filePath)
    df.to_sql(tableName, engine)
    return df


''' ----------------------------------- creating the tables needed for the code ------------------------------------ '''
createTable('C:\\Users\\OR\\Downloads\\HarryPotterCharacters.xlsx', 'Characters')
createTable('C:\\Users\\OR\\Downloads\\Harry_Potter_1_Script.xlsx', 'HarryPotter1Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\Harry_Potter_2_Script.xlsx', 'HarryPotter2Script')
createTable('C:\\Users\\OR\\Downloads\\Harry_Potter_3_Script.xlsx', 'HarryPotter3Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\\Harry_Potter_4_Script.xlsx', 'HarryPotter4Script')
createTable('C:\\Users\\OR\\PycharmProjects\\CoffeeAndHappiness\\Harry_Potter_6_Script.xlsx', 'HarryPotter6Script')


''' ----------------- finding the top 3 characters in tableName that are not Harry Hermione or Ron ----------------- '''
def topCharacters(tableName):
    topNames=engine.execute("SELECT Character, COUNT(Line) AS num_lines FROM "+tableName+" WHERE Character NOT IN ('Harry', 'Hermione', 'Ron') GROUP BY Character ORDER BY num_lines DESC LIMIT 3")
    topCharacters=pd.DataFrame(topNames, columns=['Character', 'Number of lines'])
    return topCharacters


''' ---------------------------------------------- creating a graph ------------------------------------------------ '''
def createGraph(movies, scriptsNames):
    G = nx.DiGraph()
    G.add_nodes_from(movies, bipartite=1)
    for i in range(len(scriptsNames)):
        TopCharactersTable = topCharacters(scriptsNames[i])
        #print(movies[i],": ")
        #print(TopCharactersTable)
        TopCharactersList = TopCharactersTable.Character.tolist()
        TopLinesAmount = TopCharactersTable["Number of lines"].tolist()

        # adding the top 3 characters to the graph
        G.add_nodes_from(TopCharactersList, bipartite=0)
        # adding edges between the top characters and the movie
        for j in range(len(TopCharactersList)):
            G.add_edge(TopCharactersList[j], movies[i], weight=TopLinesAmount[j])
    return G


''' ------------------------------------------ creating a two-sided graph ------------------------------------------ '''
def createTwoSidedGraph(moviesList, scriptsNames ):
    # creating a general graph
    G = createGraph(moviesList, scriptsNames)
    # painting the nodes in colors according to a condition
    G_nodes_colors = ['skyblue' if node in moviesList else 'pink' for node in G.nodes()]
    edge_labels = nx.get_edge_attributes(G, 'weight')

    # creating two sets for the bipartite graph
    U, V = bipartite.sets(G)
    pos = dict()
    pos.update((n, [1, i]) for i, n in enumerate(U))  # put nodes from U at x=1
    pos.update((n, [2, i]) for i, n in enumerate(V))  # put nodes from V at x=2

    return (G, pos, G_nodes_colors, edge_labels)


''' ----------------------------- determining edges colors according to their weight ----------------------------- '''
def edgesColorsByWeight(G):
    # coloring edges according to their weight
    huge = [(u, v) for (u, v, d) in G.edges(data=True) if d["weight"] >= 100]
    large = [(u, v) for (u, v, d) in G.edges(data=True) if d["weight"] >= 60 and d["weight"] < 100]
    small = [(u, v) for (u, v, d) in G.edges(data=True) if d["weight"] < 60]

    return (huge, large, small)



''' ---------------------------------------------------------------------------------------------------------------- '''
''' --------------------------------------------------- OUTPUTS ---------------------------------------------------- '''
''' ---------------------------------------------------------------------------------------------------------------- '''

''' ----------- creating a two-sided graph with the movies' names and the top 3 characters in each movie ----------- '''

# The graph is disconnected, therefore I create two separate graphs and then plot them together

# creating the first graph:
movies=['Harry Potter 1', 'Harry Potter 3', 'Harry Potter 4', 'Harry Potter 6']
scriptsNames=['HarryPotter1Script', 'HarryPotter3Script','HarryPotter4Script','HarryPotter6Script']
(G, pos1, G_nodes_colors, G_edge_labels)=createTwoSidedGraph(movies, scriptsNames)

# creating the second graph:
HP2MovieName='Harry Potter 2'
HP2MovieScript='HarryPotter2Script'
(HP2, pos2, HP2_nodes_colors, HP2_edge_labels)=createTwoSidedGraph([HP2MovieName], [HP2MovieScript])

# In order to display G and HP2 graph on the dame plot, shifting the x values of every node in HP2 up by 8
for k,v in pos2.items():
    v[1] += 8

# painting G's edges according to their weight (number of lines)
(huge, large, small)=edgesColorsByWeight(G)
nx.draw_networkx_edges(G, pos1, edgelist=huge, width=3, edge_color="r")
nx.draw_networkx_edges(G, pos1, edgelist=large, width=3, edge_color="y")
nx.draw_networkx_edges(G, pos1, edgelist=small, width=3, edge_color="g")

# painting HP2's edges according to their weight (number of lines)
(huge2, large2, small2)=edgesColorsByWeight(HP2)
nx.draw_networkx_edges(HP2, pos2, edgelist=huge2, width=3, edge_color="r")
nx.draw_networkx_edges(HP2, pos2, edgelist=large2, width=3, edge_color="y")
nx.draw_networkx_edges(HP2, pos2, edgelist=small2, width=3, edge_color="g")

# drawing the final graph:
nx.draw(G, pos=pos1, with_labels = True, node_color=G_nodes_colors, font_size=6, node_size=1000)
nx.draw(HP2, pos=pos2, with_labels = True, node_color=HP2_nodes_colors, font_size=6, node_size=1000)

# plotting the final graph
plt.show()