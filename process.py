import pandas

data = pandas.read_excel('Sensorium_use_4_talks.xlsx', sheet_name='talk schedule', converters={'time': str})
n = data.shape[0]

f = open('snippet.txt')
template = f.read()
f.close()

output = open('output.md', 'w')

for i in range(n):
    talk = data.iloc[i, :]
    time = talk.time
    presenter = talk.presenter
    email = talk.emails
    authors = talk['author list']
    abstract = talk.abstract
    title = talk.title

    filled = template.replace('[time]', time)
    filled = filled.replace('[presenter]', presenter)
    filled = filled.replace('[title]', title)

    filled = filled.replace('[email]', email)
    filled = filled.replace('[authors]', authors)
    filled = filled.replace('[abstract]', abstract)

    output.write(filled + '\n')

output.close()

