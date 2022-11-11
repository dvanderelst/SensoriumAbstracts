import pandas


def read_contents(filename):
    fle = open(filename)
    txt = fle.read()
    fle.close()
    return txt


#
# TALKS
#

data = pandas.read_excel('Sensorium_use_4_talks.xlsx', sheet_name='talk schedule')
n = data.shape[0]

f = open('snippet_talk.txt')
template = f.read()
f.close()

output = open('talks_output.md', 'w')

for i in range(n):
    talk = data.iloc[i, :]
    print(list(talk))
    print()
    time = talk.time
    presenter = talk.presenter
    email = talk.emails
    authors = talk['author list']
    abstract = talk.abstract
    title = talk.title

    time = str(time)[0:-3]


    filled = template.replace('[time]', time)
    filled = filled.replace('[presenter]', presenter)
    filled = filled.replace('[title]', title)

    filled = filled.replace('[email]', email)
    filled = filled.replace('[authors]', authors)
    filled = filled.replace('[abstract]', abstract)

    output.write(filled + '\n')

output.close()

#
# TALKS
#

data = pandas.read_excel('Sensorium_use_4_posters.xlsx', sheet_name='poster numbers', converters={'time': str})
data = data.sort_values(by='Session')
data['session'] = data['Session']
data['author list'] = data['authors']
n = data.shape[0]

f = open('snippet_poster.txt')
template = f.read()
f.close()

output1 = open('posters_1_output.md', 'w')
output2 = open('posters_2_output.md', 'w')
for i in range(n):
    poster = data.iloc[i, :]
    print(list(poster))
    print()
    session = str(poster.session)
    presenter = poster.presenter
    authors = poster['author list']
    abstract = poster.abstract
    title = poster.title

    if pandas.isna(authors): authors = ''

    filled = template.replace('[presenter]', presenter)
    filled = filled.replace('[title]', title)

    filled = filled.replace('[authors]', authors)
    filled = filled.replace('[abstract]', abstract)

    if session == '1': output1.write(filled + '\n')
    if session == '2': output2.write(filled + '\n')

output1.close()
output2.close()

#
#
#

talks = read_contents('talks_output.md')
session1 = read_contents('posters_1_output.md')
session2 = read_contents('posters_2_output.md')

final = open('docs/index.md', 'w')
final.write('<h1>Talks</h1>\n')
final.write(talks)
final.write('<h1>Poster Session 1</h1>\n')
final.write(session1)
final.write('<h1>Poster Session 2</h1>\n')
final.write(session2)
final.close()