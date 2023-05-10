
import csv
import json
import pandas as pd
import os

class Tools:
    def __init__(self) -> None:
        pass

    def convertXlsxToCsv(self, excelFeedackFile):
        dataFrameToBeConverted = pd.read_excel(excelFeedackFile)
        
        csvFilepath = f"../feedbacksCsv/{excelFeedackFile.replace('.xlsx', '.csv')}"
        dirname = os.path.dirname(__file__)
        csvFilepath = os.path.join(dirname, csvFilepath)

        dataFrameToBeConverted.to_csv(csvFilepath, index=None, header=True, sep=';', encoding='utf-8')
        return csvFilepath

    def convertCsvToList(self, filename):
        with open(filename, 'r', encoding='utf-8') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=';', dialect='excel')
            return list(csv_reader)

    def readJSON(self, feedbackDatabase):
        with open(feedbackDatabase, 'r') as feedbackDataBaseFile:
            return json.load(feedbackDataBaseFile)

    def replaceStrings(self, stringToChange, dictOfOldNewPairs):
        for key, value in dictOfOldNewPairs.items():
            stringToChange = stringToChange.replace(key, value)
        return stringToChange
    
    def readDB(self, feedbackDatabase):
        pass