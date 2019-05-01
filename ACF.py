# -*- coding: utf-8 -*-
"""
Created on Mon Apr 22 15:26:57 2019

@author: ykfri
"""

import numpy as np
import pandas as pd

Metric = np.array(('ROA', 'ROIC', 'ROCE', 'Gross Margin', 'Profit Margin', 
                  'EBITDA Margin', 'Op CF Margin', 'Curr Ratio',
                  'LTDE Ratio', 'Market Cap', 'BV Equity', 'Revenue', 'Cash Dividends',
                  'Net Cash Flow', 'Employees'))
Company = np.array(('AMERICAN ELECTRIC POWER CO', 'NORAM ENERGY CORP', 'CENTERIOR ENERGY CORP',
                    'COLUMBIA ENERGY GROUP', 'UNICOM CORP', 'CONSOLIDATED EDISON INC',
                    'CONSOLIDATED NATURAL GAS CO', 'DTE ENERGY CO', 'DOMINION ENERGY INC',
                    'DUKE ENERGY CORP', 'NEXTERA ENERGY INC', 'CENTERPOINT ENERGY INC', 
                    'ENRON CORP', 'NIAGARA MOHAWK HOLDINGS INC', 'NISOURCE INC', 'FIRSTENERGY CORP',
                    'PG&E CORP', 'PANENERGY CORP', 'PEOPLES ENERGY CORP', 'EXELON CORP',
                    'PUBLIC SERVICE ENTRP GRP INC', 'EDISON INTERNATIONAL', 'SOUTHERN CO', 
                    'ENERGY FUTURE HOLDINGS CORP', 'WILLIAMS COS INC', 'CLEVELAND ELECTRIC ILLUM',
                    'AES CORP', 'AMERICAN WATER WORKS CO INC'))

# Autocorrelation
ACF = np.zeros((len(Company), len(Metric)))
ACF = pd.DataFrame(ACF, index=Company, columns=Metric)

for a in range(Metric.shape[0]):
    Data = pd.read_excel('Fundamentals.xlsx', sheet_name= Metric[a])
    for k in range(1, Data.shape[1]):
        ACF.iloc[k-1, a] = np.corrcoef(Data.iloc[:, k].dropna()[:-1], Data.iloc[:, k].dropna()[1:])[0, 1]


ACF.to_excel('ACF.xlsx', index = True, header=True)

# Count number of elements
Count = np.zeros((len(Company), len(Metric)))
Count = pd.DataFrame(Count, index=Company, columns=Metric)

for a in range(Metric.shape[0]):
    Data = pd.read_excel('Fundamentals.xlsx', sheet_name= Metric[a])
    for k in range(1, Data.shape[1]):
        Count.iloc[k-1, a] = Data.iloc[:, k].count()


# T-stat of Autocorrelation
Tstat = np.zeros((len(Company), len(Metric)))
Tstat = pd.DataFrame(Tstat, index=Company, columns=Metric)

for a in range(ACF.shape[0]):
    for k in range(ACF.shape[1]):
        if np.isnan(ACF.iloc[a, k]) is True:
            Tstat.iloc[a, k] = np.nan
        else:
            Tstat.iloc[a, k] = ACF.iloc[a, k] * np.sqrt((Count.iloc[a, k]-2)/(1-ACF.iloc[a, k]**2))
         
Tstat.to_excel('Tstat.xlsx', index = True, header=True)            



