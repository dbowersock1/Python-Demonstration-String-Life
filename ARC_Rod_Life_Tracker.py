#! Will import as needed to understand what each import does
import pandas as pd
import math
from datetime import datetime

#! OVERVIEW
def main():
    #! Retreive program data from JSON file ... is there a need?

    #! Read Excel File and create Pandas Data Frame from information
    excelFileLocation = r"E:\Rod Component Analysis\15-36 Pad Rod Components.xlsx"
    df = pd.read_excel(excelFileLocation)
 
    #! Processes data from DataTable
    ProcessData(df)

    #! Store program data in JSON File so program 

    print ("test")

    # Processes passed data file and inputs specific information into arrays.
def ProcessData(df):

    # Key arrays that will be exported as data. 
    arrayLength = 262 # Each array index is a joint of rod. 262 = 262 joints = 2000m. 
    superiorRodString = [{ # Highest average joint. ie) which rod string produces the best results
                "size" : 0,
                "grade" : "default",
                "daysInHole" : 0
        }]*arrayLength     
    dataDict = dict()

    # Iterates over dataframe row x row, or rod component x rod component
    for i in range(df.index.size):
        # Pulls key information from dataframe row
        dataSeries = df.iloc[i] # iloc (Pandas) returns the values + headers of index (integer only) as a new series w/ headers as index and values as values
        startDate = dataSeries["Run Date"]
        if (pd.isnull(dataSeries["Pull Date"])): #! Catches error where there is a blank pull date ... often meaning it hasn't been pulled yet. ie) still in hole
            endDate = datetime.now()
        else:   
            endDate = dataSeries ["Pull Date"]
        daysInHole = (endDate - startDate).days
        joints = dataSeries["Joints"]
        if (dataSeries["Top Depth"] < 0 ): # Catches error where top depth is negative value
            firstJoint = 0
        else: 
            firstJoint = math.floor(dataSeries["Top Depth"] / 7.62) # Rounds down to get # of rods ... a fraction won't do
        grade = str(dataSeries["Grade"])
        size = round(dataSeries["OD Nominal"])
        rodGradeSizeTuple = (grade, size)

        # Data population

        #populate data dict
        if i == 0 :
            dataDict[rodGradeSizeTuple] = [daysInHole]
        elif rodGradeSizeTuple in dataDict.keys() :
            dataDict[rodGradeSizeTuple].append(daysInHole)
        else:
            dataDict[rodGradeSizeTuple] = [daysInHole]

        # populate SuperiorRodString
        for j in range(joints):
            #Populates arrays with values from data frame row
            element = j + firstJoint
            arrayElement = superiorRodString[element]["daysInHole"]
            if (daysInHole > arrayElement):
                superiorRodString [element] = {
                "size" : size,
                "grade" : grade,
                "daysInHole" : daysInHole
                }
           

    dfExport1 = pd.DataFrame(superiorRodString)
    dfExport1.to_excel(r"E:\Rod Component Analysis\dataExport.xlsx")
    
     


if __name__=="__main__":
    main()