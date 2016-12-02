import pandas as pd
import os


def extract_on_off(filename):
    a = pd.read_excel(filename, skiprows=31, index_col=None, na_values=['NA'])
    a.sort_values(by='Unnamed: 0', inplace=True)
    a.drop(a.columns[[1, 3, 4, 5, 6, 7, 10, 11, 12]], axis=1, inplace=True)
    a.dropna(how='any', inplace=True)
    a['Date'], a['Time'] = zip(*a['Unnamed: 0'].apply(lambda x: x.split(' ', 1)))
    b = a.groupby('Date')
    result = b.head(1).append(b.tail(1))
    result = result.sort_values(['Date','Time'])
    result.rename(columns={'Unnamed: 2': 'Door', 'Unnamed: 8': 'Access', 'Unnamed: 9': 'Cardholder',
                           'Unnamed: 13': 'Card number'}, inplace=True)
    result.drop(result.columns[[0]], axis=1, inplace=True)
    result = result[['Date', 'Time', 'Door', 'Access', 'Cardholder', 'Card number']]
    result.index = range(1, len(result) + 1)
    result.to_excel('_' + filename, sheet_name='sheet1')


def find_xls(directory='.'):
    files = []
    for f in os.listdir(directory):
        if os.path.isfile(f) and f.endswith('.xls'):
            files.append(f)
    return files


def init():
    for f in find_xls():
        extract_on_off(f)


if __name__ == '__main__':
    init()