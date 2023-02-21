import pandas as pd
import numpy as np
import seaborn as sns
from matplotlib import pyplot as plt
import os

# Three files in the path for three stock indexes
# sp500
# DJI
# nasdaq

path = r"C:\Users\wbouley\Documents\Historical_stock_CSVs/"

fileNames = os.listdir(path)
fileNames = [file for file in fileNames if '.csv' in file]

for file in fileNames:
    df = pd.read_csv(path + file, index_col=0)
    plt.plot(df)
    
plt.axis([2022, 1900, 1000, 40000])
plt.show()
