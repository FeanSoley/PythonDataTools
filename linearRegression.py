import numpy as np
from sklearn.linear_model import LinearRegression
import pandas as pd

# Get data from CSV file
data=pd.read_csv('Polution Data Messy.csv', sep=',',header=0)

# Get rid of time column
data = data.drop(columns=['Time'])
print(data.values)

yData = data[['Acetone Output']]
xData = data.drop(columns=['Acetone Output'])

print(xData)
print(yData)

reg = LinearRegression().fit(xData, yData)
print(reg.score(xData, yData))
print(reg.coef_)