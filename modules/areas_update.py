from tools import Tools
import json


class AreasUpdate(Tools):
    def __init__(self, excelAreasFeedackFile, feedbackDatabase):
        super().__init__()
        self.excelAreasFeedackFile = excelAreasFeedackFile
        self.feedbackDatabase = feedbackDatabase
        self.csvAreasFeedbackFile = None
        self.listAreasFeedbackFile = None
        self.ship = 'NB518'


    def applicableString(self, stringToChange):
        stringToChange = stringToChange.strip()
        stringToChange = stringToChange.replace('\t', ' ')
        stringToChange = stringToChange.replace('\n', ' ')
        print(stringToChange)
        return stringToChange


    def extractAreas(self, Q_A_dict):
        try:
            return Q_A_dict['§'].split(';')
        except:
            return ['']

    def updateAreasDatabase(self):

        self.initializeDatabase()

        self.feedbackDatabase = self.relativeFilepathToAbsolute(self.feedbackDatabase)
        questions = self.listAreasFeedbackFile[0]
        database = self.readJSON(self.feedbackDatabase)


        with open(self.feedbackDatabase, 'w') as databaseToWrite:
            try:
                
                databaseToUpdate = database
                
                for row in self.listAreasFeedbackFile[1:]:
                    
                    questionAnswerDict = self.questionAnswerDict(row, questions)
                    questionAnswerDict = self.removeQuestionsWithNoAnswer(questionAnswerDict)
                    areasList = self.extractAreas(questionAnswerDict)


                    for specificArea in areasList:
                        if specificArea == '':
                            continue

                        specificArea = self.applicableString(specificArea)
                        
                        if specificArea not in databaseToUpdate['feedbacks']['areas']:
                            databaseToUpdate['feedbacks']['areas'][specificArea] = {}

                        if self.ship not in databaseToUpdate['feedbacks']['areas'][specificArea]:
                            databaseToUpdate['feedbacks']['areas'][specificArea][self.ship] = {}

                        for question, answer in questionAnswerDict.items():

                            if question not in databaseToUpdate['feedbacks']['areas'][specificArea][self.ship]:
                                databaseToUpdate['feedbacks']['areas'][specificArea][self.ship][question] = []
                            
                            if answer not in databaseToUpdate['feedbacks']['areas'][specificArea][self.ship][question]:
                               databaseToUpdate['feedbacks']['areas'][specificArea][self.ship][question].append(answer)

                json.dump(databaseToUpdate, databaseToWrite, indent=4)
            except Exception as e:
                print('Error in updateDatabase()', e)
                json.dump(database, databaseToWrite, indent=4)
        databaseToWrite.close()


    def test(self):
        questions = self.listAreasFeedbackFile[0]
        qadict = self.questionAnswerDict(self.listAreasFeedbackFile[1], questions)
        print(self.removeQuestionsWithNoAnswer(qadict))
        dbPath = self.relativeFilepathToAbsolute('../databases/feedbackDatabaseTest.json')
        database = self.readJSON(dbPath)
        print(database['feedbacks']['Areas'])


    def main(self):
        self.createCsvDir()
        self.csvAreasFeedbackFile = self.convertXlsxToCsv(self.excelAreasFeedackFile)
        self.listAreasFeedbackFile = self.convertCsvToList(self.csvAreasFeedbackFile)
        #self.test()
        self.updateAreasDatabase()

if __name__ == '__main__':
    excelFilepath = "NB518_areas.xlsx"
    
    AreasUpdate = AreasUpdate(excelFilepath, '../databases/feedbackDatabaseTest.json')
    AreasUpdate.main()
