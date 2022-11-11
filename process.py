import pandas

data = pandas.read_excel('Sensorium_use_4_talks.xlsx', sheet_name='talk schedule', converters={'time': str})
n = data.shape[0]

for i in range(n):
    talk = data.iloc[i, :]
    time = talk.time
    presenter = talk.presenter
    email = talk.emails
    authors = talk['author list']
    abstract = talk.abstract
    title = talk.title

