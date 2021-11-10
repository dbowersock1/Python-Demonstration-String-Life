import pandas as pd
import math
from datetime import datetime
# Modules used to output data in csv
import csv 
from itertools import zip_longest

#! OVERVIEW
#! Maybe create graphs here in python instead of exporting to excel?? Maybe that's easier??

def main():

    # Read Excel File and create Pandas Data Frame from information
    excelFileLocation = r"E:\Rod Component Analysis\South Wells.xlsx"
    df = pd.read_excel(excelFileLocation)
 
    # Processes data from DataTable. Returns dataDictioinary. Each index corrosponds to joint of rod
    dataDict = ProcessData(df)

    #! Logic that creates the superior rod string. Returns array!
    superiorRodStringAvg, superiorRodStringHighestRun = dataAnalyze(dataDict)

    # Splits data into excel formats
    excelCreation(dataDict, superiorRodStringAvg, superiorRodStringHighestRun)

    print ("end of program")

    # Processes passed data file and inputs specific information into arrays.
def ProcessData(df):
    arrayLength = 350 # Each array index is a joint of rod. eg) 262 = 262 joints = 2000m.  
    dataDict = [] # List of Dicts of Dicts of lists. Contains all data
    for i in range(arrayLength):
        dataDict.append(dict())
   
    # DATA EXTRACTION
    # Iterates over dataframe row x row, or rod component x rod component
    for i in range(df.index.size):
        dataSeries = df.iloc[i] # iloc (Pandas) returns the values + headers of index (integer only) as a new series w/ headers as index and values as values
        startDate = dataSeries["Run Date"]
        if (pd.isnull(dataSeries["Pull Date"])): #! Catches error where there is a blank pull date ... often meaning it hasn't been pulled yet. ie) still in hole
            endDate = datetime.now()
        else:   
            endDate = dataSeries ["Pull Date"]
        daysInHole = (endDate - startDate).days
        joints = int(dataSeries["Joints"])
        if (dataSeries["Top Depth"] < 0 ): # Catches error where top depth is negative value
            firstJoint = 0
        else:
            offset = 0
            firstJoint = math.floor((dataSeries["Top Depth"]+offset) / 7.62) # Rounds down to get # of rods ... a fraction won't do
        job = dataSeries["UWI"] + " " + str(dataSeries["Run Date"])
        size = round(dataSeries["OD Nominal"])
        # Filters/Fixes rod grade to certain values
        grade = str(dataSeries["Grade"])
        if "k" in grade or "K" in grade or "d" in grade or "D" in grade:
            gradeinput = "D"
        elif "m" in grade or "M" in grade:
            gradeinput = "MMS"
        elif 'n' in grade or 'N' in grade:
            gradeinput = "N"
        elif "hs" == grade or "HS" == grade:
            gradeinput = "HS"
        else: 
            gradeinput = "Special/Alpha"
        rodGradeSizeTuple = (gradeinput, size)

        # DATA POPULATION
        for j in range(joints): 
            index = j + firstJoint
            if rodGradeSizeTuple in dataDict[index]:
                dataDict[index][rodGradeSizeTuple]["daysInHole"].append(daysInHole)
                dataDict[index][rodGradeSizeTuple]["job"].append(job)
            else:
                dataDict[index][rodGradeSizeTuple] = {"daysInHole" : [daysInHole], "job" : [job]}
    return dataDict
    

def dataAnalyze(dataDict):
    # [2] important data sets; highest avg and highest days in hole
    superiorRodStringAvg = []
    for i in range(len(dataDict)):
        superiorRodStringAvg.append({
            "gradeSize" : "default",
            "daysInHole" : 0,
            "count" : 0
            })
    superiorRodStringHighestRun = []
    for i in range(len(dataDict)):
        superiorRodStringHighestRun.append({
            "gradeSize" : "default",
            "daysInHole" : 0,
            })
    for i in range(len(dataDict)):
        for key, value in dataDict[i].items():
            # Superior Rod String Avg Updater
            avg = sum(value["daysInHole"]) / len(value["daysInHole"])
            if avg > superiorRodStringAvg[i]["daysInHole"] and len(value["daysInHole"])>5:
                superiorRodStringAvg[i]["daysInHole"] = avg
                superiorRodStringAvg[i]["gradeSize"] = key 
                superiorRodStringAvg[i]["count"] = len(value["daysInHole"])
            #! Superior Rod String Highest Run Updater
            if max(value["daysInHole"]) > superiorRodStringHighestRun[i]["daysInHole"]:
                superiorRodStringHighestRun[i]["daysInHole"] = max(value["daysInHole"]) 
                superiorRodStringHighestRun[i]["gradeSize"] = key

    return superiorRodStringAvg, superiorRodStringHighestRun

def excelCreation(dataDict, superiorRodStringAvg, superiorRodStringHighestRun):
    # Superior Rod String Export
    df = pd.DataFrame(superiorRodStringAvg)
    df2 = pd.DataFrame(superiorRodStringHighestRun)
    exportLink = r'E:\Rod Component Analysis\Superior Rod String Avg.xlsx'
    exportLink2 = r'E:\Rod Component Analysis\Superior Rod String Highest Run.xlsx'
    df.to_excel(exportLink)
    df2.to_excel(exportLink2)

    # Supportig Data Export
    with open (r'E:\Rod Component Analysis\All Data.csv', 'w', newline="") as f:
         writer = csv.writer(f)
         writer.writerow(["Joint", "Grade/Size", "Days In Hole", "Job"])
         for i in range(len(dataDict)):
             for key in dataDict[i].keys():
                 for j in range(len(dataDict[i][key]["daysInHole"])):
                       writer.writerow([i,key,dataDict[i][key]["daysInHole"][j],dataDict[i][key]["job"][j]])
       
if __name__=="__main__":
    main()