#Fundamental Factor Model - Observe stock exposures to factors... estimate factor , Cross Sectional
#Economic Factor Model- Time Series Model Observe Returns to each factor  and estimate factor loadings
import numpy as np
import pandas as pd
import math as math
import openpyxl
import numpy as np
from scipy.stats.mstats import winsorize # we are on python 2.7 since winsorize doesnt exist in 3.6
import matplotlib.pyplot as plt





xlsx = pd.ExcelFile('ffm.xlsx')
df = pd.read_excel(xlsx, '2013')  # for now just do it to a sheet at a time.. later on loop it for sheets.
df.set_index('Ticker', inplace=True)
#data frame starts from the 7th
df['logcap'] = df['mkt value']  #kinda just a placeholder? can be more efficent


def winsorize_data():#winsorizing using this method gives a slightly different value than excel
    for mktvalue in range(len(df['mkt value'])):
        df['logcap'][mktvalue] = math.log(df['mkt value'][mktvalue],10)
        mktvalue += 1

    logcap = df['logcap'].values # turned column into a list
    win_logcap_list = winsorize(logcap, limits =(.01, .01)) # winsorized off data 1% on each side.
    # can do a loop and rename those column title by hand at the end by renaming 'old title+winsorize'
    df['winsorizedlogcap'] = win_logcap_list  #column winsorizedlogcap is my list of logcap
    df['winsorizedgpm'] = winsorize(df['Annual GPM'].values, limits=(.01, .01))
    df['winsorizedbp'] = winsorize(df['Annual BP'].values, limits =(.01, .01))
    df['winsorizedepsg'] = winsorize(df['EPS Growth(1y)'].values, limits=(.01, .01))
    # if statement, if exists skip it instead of keep rewriting

    #writer = pd.ExcelWriter('ImportedData.xlsx')
    #df.to_excel(writer,'Sheet1')
    #print (df)

def standardize_data():
    meanlist = []
    stdevlist = []
    df = pd.read_excel(pd.ExcelFile('ImportedData.xlsx'))

    print(df.columns[7])

    for column in range(0,4):
        mean = np.mean(df.ix[:,7+column].values)
        meanlist.append(mean)

        stdev = np.std(df.ix[:,7+column].values)
        stdevlist.append(stdev)

    for value in range(0,4):
        df['z{}'.format(df.columns[7+value])] = (df[df.columns[7+value]]-meanlist[value])/stdevlist[value]

    #use some if statement
    df.to_excel('ImportedData.xlsx','Sheet1')



def trim_data():
    df = pd.read_excel(pd.ExcelFile('ImportedData.xlsx'))

    for trimcolumn in range(0,4):
        zlist = []

        for zscore in df.ix[:,11+trimcolumn]:
            if zscore > 3:
                zlist.append(3)
            elif zscore < -3:
                zlist.append(-3)
            else:
                zlist.append(zscore)
        df['trimmed{}'.format(df.columns[7 + trimcolumn])] = zlist

        print zlist
    df.to_excel('ImportedData.xlsx', 'Sheet1')

#Need to regress factors vs return
def regression_analysis():
    df = pd.read_excel(pd.ExcelFile('ImportedData.xlsx'))
    df.set_index('Ticker', inplace=True)
    print df.head()



    #graph part
    fig, axs = plt.subplots(1, 4, sharey=True)
    df.plot(kind='scatter', x='trimmedwinsorizedlogcap', y='Return', ax=axs[0], figsize=(16, 8))
    df.plot(kind='scatter', x='trimmedwinsorizedgpm', y='Return', ax=axs[1])
    df.plot(kind='scatter', x='trimmedwinsorizedbp', y='Return', ax=axs[2])
    df.plot(kind='scatter', x='trimmedwinsorizedepsg', y='Return', ax=axs[3])
regression_analysis()


