import pandas



#
# TALKS
#

data = pandas.read_excel('Sensorium_use_4_talks.xlsx', sheet_name='talk schedule', converters={'time': str})
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

    time = time.rstrip(':00')

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


output = open('posters_output.md', 'w')

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

    filled = template.replace('[session]', session)
    filled = filled.replace('[presenter]', presenter)
    filled = filled.replace('[title]', title)

    filled = filled.replace('[authors]', authors)
    filled = filled.replace('[abstract]', abstract)

    output.write(filled + '\n')

output.close()


