

import os
import pandas as pd

def delfile(file):
    # removes a file if it exists 
    if os.path.exists(file):
        os.remove(file)

def func(x):
    if x.iloc[23] != '--':
        return x.iloc[23]
    return x.iloc[22]


def funcswitch(x):
    if (pd.isnull(x.iloc[23])) | (x.iloc[23] == '--'):        
        return x.iloc[22]
    return x.iloc[23]      
  

