import pandas as pd
import csv
import json

class FeedbackDataBase:
    def __init__(self, excelFeedackFiles, feedbackDatabase='database.json', mode=None):
        self.excelFeedackFile1 = excelFeedackFiles[0]
        self.excelFeedackFile2 = excelFeedackFiles[1]
        self.csvFeedbackFile1 = excelFeedackFiles[0].replace('.xlsx', '.csv')
        self.csvFeedbackFile2 = excelFeedackFiles[1].replace('.xlsx', '.csv')
        self.listFeedbackFile1 = None
        self.listFeedbackFile2 = None
        self.feedbackDatabase = feedbackDatabase
        self.mode = mode



    def convertXlsxToCsv(self, excelFeedackFile):
        dataFrameToBeConverted = pd.read_excel(excelFeedackFile)
        dataFrameToBeConverted.to_csv(excelFeedackFile.replace('.xlsx', '.csv'), index=None, header=True, sep=';', encoding='utf-8')
        return True

    def convertCsvToList(self, filename):
        with open(filename, 'r', encoding='utf-8') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=';', dialect='excel')
            return list(csv_reader)

    def readJSON(self):
        with open(self.feedbackDatabase, 'r') as feedbackDataBaseFile:
            return json.load(feedbackDataBaseFile)

    def replaceStrings(self, stringToChange, dictOfOldNewPairs):
        for key, value in dictOfOldNewPairs.items():
            stringToChange = stringToChange.replace(key, value)
        return stringToChange

    def updateDatabaseAreaOutfitting(self):
        
        self.mode='Area outfitting'
        questions = self.listFeedbackFile1[0]
        database = self.readJSON()
        
        
        with open(self.feedbackDatabase, 'w') as databaseToWrite:
            try:
                databaseToUpdate = database
                for row in self.listFeedbackFile1[1:]:
                    areaOfResponsibility = self.replaceStrings(row[5], {'\t': ' '})
                    if areaOfResponsibility not in databaseToUpdate['NB518'][self.mode]:
                        databaseToUpdate['NB518'][self.mode][areaOfResponsibility] = {}
                    for question, answer in zip(questions, row):                        
                        if question not in databaseToUpdate['NB518'][self.mode][areaOfResponsibility]:
                            databaseToUpdate['NB518'][self.mode][areaOfResponsibility][question] = [answer]
                        if answer not in databaseToUpdate['NB518'][self.mode][areaOfResponsibility][question]:
                            databaseToUpdate['NB518'][self.mode][areaOfResponsibility][question].append(answer)
                        
                json.dump(databaseToUpdate, databaseToWrite, indent=4)
            except Exception as e:
                print('Error in updateDatabase()', e)
                json.dump(database, databaseToWrite, indent=4)
        databaseToWrite.close()


    def main(self):
        self.convertXlsxToCsv(self.excelFeedackFile1) # works
        self.convertXlsxToCsv(self.excelFeedackFile2) # works
        self.listFeedbackFile1 = self.convertCsvToList(self.csvFeedbackFile1) # works
        self.listFeedbackFile2 = self.convertCsvToList(self.csvFeedbackFile2) # works
        self.updateDatabaseAreaOutfitting() # pending

if __name__ == '__main__':
    #feedbackFile = str(input('Syötä palaute tiedosto: '))
    feedbackFile1 = 'NB518_e1.xlsx'
    feedbackFile2 = 'NB518_e2.xlsx'
    feedbackDataBase = FeedbackDataBase([feedbackFile1, feedbackFile2], 'database1.json')
    feedbackDataBase.main()