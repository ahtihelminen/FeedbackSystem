from tools import Tools
import json
import os

class SystemsUpdate(Tools):
    def __init__(self, excelSystemsFeedackFile, feedbackDatabase):
        super().__init__()
        self.excelSystemsFeedackFile = excelSystemsFeedackFile
        self.feedbackDatabase = feedbackDatabase
        self.csvSystemsFeedbackFile = None
        self.listSystemsFeedbackFile = None
        self.ship = 'NB518'
    

    

    def extractSystems(self, Q_A_dict):
        try:
            systemCode = Q_A_dict['valitse littera']
        except:
            try:
                systemCode = Q_A_dict['choose system code']
            except Exception as e:
                print('Error in extractSystems()', e)
                return False
        finally:
            for question, answer in Q_A_dict.items():
                try:
                    if question.split(' ')[0] == systemCode:
                        return systemCode, answer.split(';')
                except:
                    pass


    def applicableString(self, stringToChange):
        stringToChange = stringToChange.strip()
        stringToChange = stringToChange.replace('\t', ' ')
        stringToChange = stringToChange.replace('\n', ' ')
        return stringToChange



    def updateSystemsDatabase(self):

        self.feedbackDatabase = self.relativeFilepathToAbsolute(self.feedbackDatabase)
        questions = self.listSystemsFeedbackFile[0]
        database = self.readJSON(self.feedbackDatabase)


        with open(self.feedbackDatabase, 'w') as databaseToWrite:
            try:
                
                databaseToUpdate = database
                
                for row in self.listSystemsFeedbackFile[1:]:
                    
                    questionAnswerDict = self.questionAnswerDict(row, questions)
                    questionAnswerDict = self.removeQuestionsWithNoAnswer(questionAnswerDict)
                    systemCode, specificSystems = self.extractSystems(questionAnswerDict)
                    
                    for specificSystem in specificSystems:
                        if specificSystem == '':
                            continue

                        specificSystem = self.applicableString(specificSystem)
                        
                        if specificSystem not in databaseToUpdate['feedbacks']['systems'][systemCode]:
                            databaseToUpdate['feedbacks']['systems'][systemCode][specificSystem] = {}

                        if self.ship not in databaseToUpdate['feedbacks']['systems'][systemCode][specificSystem]:
                            databaseToUpdate['feedbacks']['systems'][systemCode][specificSystem][self.ship] = {}

                        for question, answer in questionAnswerDict.items():

                            if question not in databaseToUpdate['feedbacks']['systems'][systemCode][specificSystem][self.ship]:
                                databaseToUpdate['feedbacks']['systems'][systemCode][specificSystem][self.ship][question] = []
                            
                            if answer not in databaseToUpdate['feedbacks']['systems'][systemCode][specificSystem][self.ship][question]:
                                databaseToUpdate['feedbacks']['systems'][systemCode][specificSystem][self.ship][question].append(answer)

                json.dump(databaseToUpdate, databaseToWrite, indent=4)
            except Exception as e:
                print('Error in updateDatabase()', e)
                json.dump(database, databaseToWrite, indent=4)
        databaseToWrite.close()




    def test(self):
        questions = self.listSystemsFeedbackFile[0]
        qadict = self.questionAnswerDict(self.listSystemsFeedbackFile[1], questions)
        print(self.removeQuestionsWithNoAnswer(qadict))
        print(self.extractSystems(qadict))
        dbPath = self.relativeFilepathToAbsolute('../databases/feedbackDatabase.json')
        database = self.readJSON(dbPath)
        print(database['feedbacks']['systems'])

    def main(self):
        self.csvSystemsFeedbackFile = self.convertXlsxToCsv(self.excelSystemsFeedackFile)
        self.listSystemsFeedbackFile = self.convertCsvToList(self.csvSystemsFeedbackFile)
        self.test()
        self.updateSystemsDatabase()

if __name__ == '__main__':
    excelFilepath = "NB518_e2.xlsx"
    
    systemsUpdate = SystemsUpdate(excelFilepath, '../databases/feedbackDatabase.json')
    systemsUpdate.main()
