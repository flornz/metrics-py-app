from calendar import c
from cmath import nan
import pandas as pd
import numpy as np
import math

# Sheetname to read from and the Out Of Date (Sundays of the week)
sheetName = "SEP 12 - SEP 16"
outOfDate = "09/11/2022"

# Read from the Excel File (declaration of the Excel File Object)
xlsx = pd.ExcelFile("GFD - Master Weekly Metrics _ F2022.xlsx")

# getMonthlyAsk (sheetName: string, numCharities: int)
# Reads the column "B" to get all products for each corresponding charity
# Returns pandas dataframe for monthly ask
def getMonthlyAsk(sheetName, numCharities):
    mnthlyAsk = pd.read_excel(xlsx, sheetName, index_col=None, na_values=['NA'], usecols="B", skiprows=4, nrows=numCharities)
    return mnthlyAsk

# getClients(sheetName: string)
# Reads the column "A" to get all charities in the column in alphabetical order.
# Returns the array of charities. Each element is a string.
def getClients(sheetName):
    num_charities = 0
    clients = pd.read_excel(xlsx, sheetName, index_col=None, na_values=['NA'], usecols="A", skiprows=4, nrows=1000)
    client = pd.DataFrame({"Charities" : clients['GFD'].copy()}).to_numpy()
    charities = []
    for x in range(len(client)):
        if isinstance(client[x][0], str):
            num_charities += 1
            charities.append(client[x][0])
        elif math.isnan(client[x][0]):
            break
    
    return charities

# getGrossMetrics(numCharities: string, col: string)
# Reads the column 'col' and returns the dataframe. 
# Used in compileDayStats to read all columns and compile all day stats for that week.
# Returns dataframe
def getGrossMetrics(numCharities, col):
    metric = pd.read_excel(xlsx, sheetName, index_col=None, na_values=['NA'], usecols=col, skiprows=3, nrows=numCharities+1)
    gross = metric.loc[range(1, numCharities+1)]
    #newTemp.to_excel("test.xlsx")
    return gross

# def getIndexDate(date):
#     indexSeries = pd.Series({"Payment Day" : date})
#     return indexSeries

# getNetMetrics(numCharities: string, col: string)
# Similar to getGrossMetrics, reads and returns the data in column 'col'
# Used in compileDaySTats which maps through column of arrays that hold the data.
# Returns dataframe
def getNetMetrics(numCharities, col):
    metric = pd.read_excel(xlsx, sheetName, index_col=None, na_values=['NA'], usecols=col, skiprows=numCharities+26, nrows=numCharities+1)
    net = metric.loc[range(1, numCharities+1)]
    return net

# compileDayStats(data: strings[])
# compileDayStats usese the above functions to create arrays of data so they can be
#    compiled into arrays and used to create a dataframe to be written into the Excel sheet.
def compileDayStats(date):
    # Columns to read from for net + gross stats
    cols = ['J', 'P', 'V', 'AB', 'AH', 'AN']
    # gets all charities in column A (reads col 'A')
    currClients = getClients(sheetName)
    grossStatsData = []
    netStatsData = []
    dateIndex = []
    # for loop to append data into one array to create a dataframe
    for z in range(len(date)):
        tempDate = []
        for m in range(len(currClients)):
            tempDate.append(date[z])
        dateIndex.append(tempDate)
        
    # returns the monthly ask and converts into a numpy array.
    products = getMonthlyAsk(sheetName, len(currClients)).to_numpy()
    newProds = []
    # takes the numpy array and converts into a 1d array 
    # (numpy arr: [[0], [1], etc.])
    # (1d array: [0, 1, etc,])
    for prod in range(len(products)):
        newProds.append(int(products[prod][0]))

    # similar to gtoss metrics and net metrics, converts the numpy array into 1d array
    for x in range(len(cols)):
        grossStatsData.append(getGrossMetrics(len(currClients), cols[x]).to_numpy())
        netStatsData.append(getNetMetrics(len(currClients), cols[x]).to_numpy())
        tempG = []
        tempN = []
        for i in range(len(grossStatsData[-1])):
            tempG.append(int(grossStatsData[-1][i]))
            tempN.append(int(netStatsData[-1][i]))
        grossStatsData[x] = tempG
        netStatsData[x] = tempN

    # print(len(dateIndex))
    # print(len(newProds))
    # print(len(netStatsData))
    # print(len(grossStatsData))
    # print(len(currClients))
    # print(netStatsData[-1])
    #
    # initialize a datafram where all data will be appended into
    WeeklyData = pd.DataFrame()

    # create a temo dataframe to add each day of the week into the main WeeklyData
    for y in range(len(grossStatsData)):
        newInsert = pd.DataFrame({"Payment Day":dateIndex[y], 'Charity Code':currClients, 'Product':newProds, 'Gross Count':grossStatsData[y], 'Net Count':netStatsData[y]})
        WeeklyData = WeeklyData.append(newInsert)
        
    # Create date for sunday and into an array
    outofdatelist = []
    for k in range(len(currClients)):
        outofdatelist.append(outOfDate)
    
    # Grab the data for outOfDate (sundays that only have NET data)
    outOfDateNetCount = []
    outOfDateGrossCount = []
    oodnt = pd.read_excel(xlsx, sheetName, index_col=None, na_values=['NA'], usecols="AT", skiprows=len(currClients)+27, nrows=len(currClients)).to_numpy()
    for counts in range(len(oodnt)):
        outOfDateNetCount.append(int(oodnt[-1][0]))
        outOfDateGrossCount.append(0)

    # print(oodnt)
    # print(len(outofdatelist))
    # print(len(currClients))
    # print(len(newProds))
    # print(len(outOfDateGrossCount))
    # print(len(outOfDateNetCount))

    # Create a dataframe to append onto the WeeklyData as well
    sundayInsert = pd.DataFrame({"Payment Day":outofdatelist, 'Charity Code':currClients, 'Product':newProds, 'Gross Count':outOfDateGrossCount, 'Net Count':outOfDateNetCount})
    WeeklyData = WeeklyData.append(sundayInsert)

    # Write that WeeklyData dataframe object into the excel sheet, and save it.
    WeeklyData.to_excel('newFile.xlsx', index=False)
    # WeeklyData.append_df_to_excel('Donor Count Summary.xlsx',sheet_name='Donor Count Summary',index=False)


# Main function that calls the compileDayStats function
def main():
    # Change array of dates in order to generate the right index values
    compileDayStats(["09/05/2022", "09/06/2022", "09/07/2022", "09/08/2022", "09/09/2022", "09/10/2022"])

main()

