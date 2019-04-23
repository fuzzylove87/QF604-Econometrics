# -*- coding: utf-8 -*-
"""
Created on Mon Apr 22 15:26:57 2019

@author: ykfri
"""

import numpy as np
import pandas as pd

Metric = np.array(('ROA', 'ROIC', 'ROCE', 'Gross Margin', 'Profit Margin', 
                  'EBITDA Margin', 'Op CF Margin', 'Curr Ratio', 'D E Ratio',
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

DF = np.zeros((len(Company), len(Metric)))
DF = pd.DataFrame(DF, index=Company, columns=Metric)


for a in range(Metric.shape[0]):
    Data = pd.read_excel('Fundamentals.xlsx', sheet_name= Metric[a])
    for k in range(1, Data.shape[1]):
        DF.iloc[k-1, a] = np.corrcoef(Data.iloc[:, k].dropna()[:-1], Data.iloc[:, k].dropna()[1:])[0, 1]


DF.to_excel('ACF.xlsx', index = True, header=True)



