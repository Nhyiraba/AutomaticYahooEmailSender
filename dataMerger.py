# -*- coding: utf-8 -*-
"""
Created on Mon May 11 19:02:59 2020

@author: PNDT
"""

import pandas as pd

def removeSpaceAtEndBegin(dataSet):
    
    for sp in range(len(dataSet.columns)-1):
        
        #print(dataSet.columns[sp])
        dataSet[str(dataSet.columns[sp])] = dataSet[dataSet.columns[sp]].str.strip()
        
    return dataSet
