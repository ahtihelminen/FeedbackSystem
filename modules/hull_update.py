from tools import Tools
import json


class HullUpdate(Tools):
    def __init__(self, excelHullFeedackFile, feedbackDatabase):
        super().__init__()
        self.excelHullFeedackFile = excelHullFeedackFile
        self.feedbackDatabase = feedbackDatabase
        self.csvHullFeedbackFile = None
        self.listHullFeedbackFile = None
        self.ship = 'NB518'


    def applicableString(self, stringToChange):
        stringToChange = stringToChange.strip()
        stringToChange = stringToChange.replace('\t', ' ')
        stringToChange = stringToChange.replace('\n', ' ')
        return stringToChange


    def updateHullDatabase(self):

        self.feedbackDatabase = self.relativeFilepathToAbsolute(self.feedbackDatabase)
        questions = self.listHullFeedbackFile[0]
        database = self.readJSON(self.feedbackDatabase)


        with open(self.feedbackDatabase, 'w') as databaseToWrite:
            try:
                
                databaseToUpdate = database
                
                for row in self.listHullFeedbackFile[1:]:
                    
                    questionAnswerDict = self.questionAnswerDict(row, questions)
                    questionAnswerDict = self.removeQuestionsWithNoAnswer(questionAnswerDict)
                    area = questionAnswerDict['choose your area']
                    
                    if area not in databaseToUpdate['feedbacks']['hull']:
                        databaseToUpdate['feedbacks']['hull'][area] = {}

                    if self.ship not in databaseToUpdate['feedbacks']['hull'][area]:
                        databaseToUpdate['feedbacks']['hull'][area][self.ship] = {}

                    for question, answer in questionAnswerDict.items():

                        if question not in databaseToUpdate['feedbacks']['hull'][area][self.ship]:
                            databaseToUpdate['feedbacks']['hull'][area][self.ship][question] = []
                        
                        if answer not in databaseToUpdate['feedbacks']['hull'][area][self.ship][question]:
                            databaseToUpdate['feedbacks']['hull'][area][self.ship][question].append(answer)

                json.dump(databaseToUpdate, databaseToWrite, indent=4)
            except Exception as e:
                print('Error in updateDatabase()', e)
                json.dump(database, databaseToWrite, indent=4)
        databaseToWrite.close()


    def test(self):
        questions = self.listHullFeedbackFile[0]
        qadict = self.questionAnswerDict(self.listHullFeedbackFile[1], questions)
        print(self.removeQuestionsWithNoAnswer(qadict))
        print(self.extracthull(qadict))
        dbPath = self.relativeFilepathToAbsolute('../databases/feedbackDatabaseTest.json')
        database = self.readJSON(dbPath)
        print(database['feedbacks']['hull'])


    def main(self):
        self.csvHullFeedbackFile = self.convertXlsxToCsv(self.excelHullFeedackFile)
        self.listHullFeedbackFile = self.convertCsvToList(self.csvHullFeedbackFile)
        #self.test()
        self.updateHullDatabase()

if __name__ == '__main__':
    excelFilepath = "NB518_hull.xlsx"
    
    hullUpdate = HullUpdate(excelFilepath, '../databases/feedbackDatabaseTest.json')
    hullUpdate.main()
