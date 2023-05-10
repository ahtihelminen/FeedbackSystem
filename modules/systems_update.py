from tools import Tools
import json

class SystemsUpdate(Tools):
    def __init__(self, excelSystemsFeedackFile, feedbackDatabase):
        super().__init__()
        self.excelSystemsFeedackFile = excelSystemsFeedackFile
        self.feedbackDatabase = feedbackDatabase
        self.csvSystemsFeedbackFile = None
        self.listSystemsFeedbackFile = None
    
    def updateSystemsDatabase(self):
        questions = self.listSystemsFeedbackFile[0]
        database = self.readJSON(self.feedbackDatabase)
        
        with open(self.feedbackDatabase, 'w') as databaseToWrite:
            try:
                databaseToUpdate = database
                for row in self.listSystemsFeedbackFile[1:]:
                    generalSystem = row[5] # 1000 Ship general design, 3000 Hull, ...
                    
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
        self.csvSystemsFeedbackFile = self.convertXlsxToCsv(self.excelSystemsFeedackFile)
        self.listSystemsFeedbackFile = self.convertCsvToList(self.csvSystemsFeedbackFile)
        self.updateSystemsDatabase()

if __name__ == '__main__':
    systemsUpdate = SystemsUpdate('NB518_e2.xlsx', 'database.json')
    systemsUpdate.main()
