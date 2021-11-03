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
    dataTracker =[[] for i in range(arrayLength)] # Keeps track of all the information used to create superiorRodString # doing this = [[]]*4 gives error. List of list of dictionaries. Each parent element is rod string

    # Generates unique list of rod size to grade tuples. Used later to have an average run time assosiation. Used to populate key SuperiorRodString Array!!
    rodSizes = (round(num) for num in df["OD Nominal"].tolist()) # Returns all rod sizes from dataframe/excel file
    #! Ensure rodGrades only returns a list!
    rodGrades = df["Grade"].tolist() # Returns all rod grades from dataframe/excel file
    for i in range(len(rodGrades)):
        rodGrades[i] = str(rodGrades[i])
    uniqueSizeandGradeList = list(set(zip(rodSizes, rodGrades))) # Elimnates duplicates then converts back to list
    uniqueListSizeGrade = [] # List containing all class objects of uniquRodGradeSize
    for item in uniqueSizeandGradeList:
        uniqueListSizeGrade.append(uniqueRodGradeSize(item))

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

        # Data population
        for j in range(joints):
            #! Create logic that populates unique grade and size, calculates average, and populates
            for item in uniqueListSizeGrade:
               if (grade == item.rodGrade and size == item.rodSize):
                   item.dataCollectionList.append({
                       "daysInHole" : daysInHole,
                       "UWI" : dataSeries["UWI"],
                       "Run Job" : dataSeries["Run Job"],
                       "Pull Job" : dataSeries["Pull Job"],
                       "Pull Reason" : dataSeries["Pull Reason"]
                       })
                   item.calculateAverageRunLife()
            #Populates arrays with values from data frame row
            element = j + firstJoint
            arrayElement = superiorRodString[element]["daysInHole"]
            if (daysInHole > arrayElement):
                superiorRodString [element] = {
                "size" : size,
                "grade" : grade,
                "daysInHole" : daysInHole
                }
            dataTracker[element].append({
                "size" : round(dataSeries["OD Nominal"]),
                "grade" : dataSeries["Grade"],
                "daysInHole" : daysInHole,
                "UWI" : dataSeries["UWI"],
                "Run Job" : dataSeries["Run Job"],
                "Pull Job" : dataSeries["Pull Job"],
                "Pull Reason" : dataSeries["Pull Reason"]
                })

    dfExport1 = pd.DataFrame(superiorRodString)
    dfExport2 = pd.DataFrame(dataTracker)
    dfExport1.to_excel(r"E:\Rod Component Analysis\dataExport.xlsx")
    dfExport2.to_excel(r"E:\Rod Component Analysis\dataExport2.xlsx")
     
    
class uniqueRodGradeSize:
    def __init__(self, uniqueTupleList):
        self.rodSize = uniqueTupleList[0]
        self.rodGrade = uniqueTupleList[1]
        self.averageRunLife = 0
        self.dataCollectionList = [] # elements inside defined in proccessing data list!
                
    def calculateAverageRunLife(self):
        if (self.dataCollectionList.count == 0):
            raise ValueError ("Tried calculating average when there was nothing in list!")
        count = 0
        sum = 0
        for item in self.dataCollectionList:
            count += 1
            sum += item["daysInHole"]
        self.averageRunLife = sum / count
        
        


if __name__=="__main__":
    main()