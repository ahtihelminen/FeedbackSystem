import pandas as pd
import csv
import json

class FeedbackDataBase:
    def __init__(self, excelFeedackFile, feedbackDatabase='database.json', mode=None):
        self.excelFeedackFile = excelFeedackFile
        self.csvFeedbackFile = excelFeedackFile.replace('.xlsx', '.csv')
        self.listFeedbackFile = None
        self.feedbackDatabase = feedbackDatabase
        self.mode = mode


    def convertXlsxToCsv(self, excelFeedackFile):
        dataFrameToBeConverted = pd.read_excel(excelFeedackFile)
        print(dataFrameToBeConverted)
        dataFrameToBeConverted.to_csv(self.csvFeedbackFile, index=None, header=True, sep=';', encoding='utf-8')
        return True

    def convertCsvToList(self, filename):
        with open(filename, 'r', encoding='utf-8') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=';', dialect='excel')
            return list(csv_reader)

    def updateDatabase(self):
        questions = self.listFeedbackFile[0]

        if self.mode == 'Area outfitting':
            for row in self.listFeedbackFile[1:]:
                databaseToUpdate = json.load(self.feedbackDatabase)
                for question, answer in zip(questions, row):
                    if question in databaseToUpdate:
                        databaseToUpdate['NB518'][self.mode][question].append(answer)
                    else:
                        databaseToUpdate[question] = [answer]
                json.dump(databaseToUpdate, self.feedbackDatabase, indent=4)

    def main(self):
        self.convertXlsxToCsv(self.excelFeedackFile) # works
        self.listFeedbackFile = self.convertCsvToList(self.csvFeedbackFile) # works
        self.updateDatabase() # works

if __name__ == '__main__':
    #feedbackFile = str(input('Syötä palaute tiedosto: '))
    feedbackFile = 'NB518_e1.xlsx'
    mode = 'Area outfitting'
    feedbackDataBase = FeedbackDataBase(feedbackFile, 'database.json', mode)
    feedbackDataBase.main()