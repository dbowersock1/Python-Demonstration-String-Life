import pandas as pd
import math
from datetime import datetime
# Modules used to output data in csv
import csv 
from itertools import zip_longest

#! OVERVIEW


def main():

    # Read Excel File and create Pandas Data Frame from information
    excelFileLocation = r"E:\Rod Component Analysis\15-36 Pad Rod Components.xlsx"
    df = pd.read_excel(excelFileLocation)
 
    # Processes data from DataTable. Prints to excel files
    ProcessData(df)

    print ("test")

    # Processes passed data file and inputs specific information into arrays.
def ProcessData(df):

    # Key arrays that will be exported as data. 
    arrayLength = 262 # Each array index is a joint of rod. 262 = 262 joints = 2000m. 
    superiorRodString = [{ # Highest average joint. ie) which rod string produces the best results
                "size" : 0,
                "grade" : "default",
                "AvgdaysInHole" : 0
        }]*arrayLength     
    dataDict = dict()
    dataDictJobCorr = dict()
    dataDictJointCorr = dict()

    # DATA EXTRACTION
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
        jointRange = "From: " + str(firstJoint) + " To: " + str(firstJoint+joints)
        job = dataSeries["UWI"] + " " + str(dataSeries["Run Date"])
        grade = str(dataSeries["Grade"])
        size = round(dataSeries["OD Nominal"])
        rodGradeSizeTuple = (grade, size)

        # DATA POPULATION
        # populate data dict
        if rodGradeSizeTuple in dataDict.keys() :
            dataDict[rodGradeSizeTuple].append(daysInHole)
            dataDictJobCorr[rodGradeSizeTuple].append(job)
            dataDictJointCorr[rodGradeSizeTuple].append(jointRange)
        else:
            dataDict[rodGradeSizeTuple] = [daysInHole]
            dataDictJobCorr[rodGradeSizeTuple] = [job]
            dataDictJointCorr[rodGradeSizeTuple] = [jointRange]
    

        # populate SuperiorRodString
        #! This might have to be done once all the rod grade is populated!! Otherwise the averages start to not make sense!
        
        for j in range(joints):
            #Populates arrays with values from data frame row
            element = j + firstJoint
            avgDaysInHole = sum(dataDict[rodGradeSizeTuple]) / len(dataDict[rodGradeSizeTuple])
            arrayElement = superiorRodString[element]["AvgdaysInHole"]
            if (avgDaysInHole > arrayElement):
                superiorRodString [element] = {
                "size" : size,
                "grade" : grade,
                "AvgdaysInHole" : avgDaysInHole
                }
           
    # Exports primary data. Superior rod string. Via pandas 
    dfExport1 = pd.DataFrame(superiorRodString)
    with pd.ExcelWriter(r"E:\Rod Component Analysis\SuperiorRodString.xlsx") as writer:
        dfExport1.to_excel(writer, sheet_name="Superior rod string")
    # Exports data that only can be exported to CSVs (uneven lists that cannot be exported by Pandas)
    with open ("E:\Rod Component Analysis\RodGradeAverageDaysData.csv", 'w', newline = "") as outputcsv:
        writer = csv.writer(outputcsv)
        writer.writerow(dataDict.keys())
        writer.writerows(zip_longest(*dataDict.values()))
    with open ("E:\Rod Component Analysis\RodGradeSizeJobCorr.csv", 'w', newline = "") as outputcsv:
        writer = csv.writer(outputcsv)
        writer.writerow(dataDictJobCorr.keys())
        writer.writerows(zip_longest(*dataDictJobCorr.values()))
    with open ("E:\Rod Component Analysis\RodGradeSizeJointCorr.csv", 'w', newline = "") as outputcsv:
        writer = csv.writer(outputcsv)
        writer.writerow(dataDictJointCorr.keys())
        writer.writerows(zip_longest(*dataDictJointCorr.values()))

if __name__=="__main__":
    main()