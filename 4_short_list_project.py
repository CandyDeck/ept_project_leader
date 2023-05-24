import pandas as pd
import nltk
from nltk.util import ngrams

from datetime import datetime
import os


today = datetime.now()

project = pd.read_excel(os.getcwd() + "\\" + str(today.strftime('%Y%m%d'))+ "\\" + str(today.strftime('%Y%m%d')) + '_list_project.xlsx','WORKORDERS')

result = project.groupby('Prosjektnr.')['Navn på arbeidsordre'].agg(lambda x: ''.join(x)).reset_index()



for i in result.index:
    #print(i, len(result.loc[i,'Navn på arbeidsordre']))
    lenght = range(2,len(result.loc[i,'Navn på arbeidsordre']))
    ngrams_count = {}
    for n in lenght:
        ngrams = tuple(nltk.ngrams(result.loc[i,'Navn på arbeidsordre'].split(' '),n=n))
        ngrams_count.update({' '.join(i) : ngrams.count(i) for i in ngrams})
    df = pd.DataFrame(list(zip(ngrams_count, ngrams_count.values())), columns=['Ngramm', 'Count']).sort_values(['Count'], ascending=False)
    print(i,df)

