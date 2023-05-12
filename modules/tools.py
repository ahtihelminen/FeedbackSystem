
import csv
import json
import pandas as pd
import os
from docx import Document
from docx.shared import Inches


class Tools:
    def __init__(self) -> None:
        pass

    def convertXlsxToCsv(self, excelFeedbackFile):
        
        dataFrameToBeConverted = pd.read_excel(excelFeedbackFile)
        
        filename = excelFeedbackFile.split('/')[-1]

        csvFilepathRel = f"../feedbacksCsv/{filename.replace('.xlsx', '.csv')}"
        csvFilepathAbs = self.relativeFilepathToAbsolute(csvFilepathRel)

        dataFrameToBeConverted.to_csv(csvFilepathAbs, index=None, header=True, sep=';', encoding='utf-8')
        return csvFilepathAbs

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
    
    def questionAnswerDict(self, answerRow, questions):
        answerDict = {}
        for question, answer in zip(questions, answerRow):
            answerDict[question] = answer
        return answerDict
    
    def relativeFilepathToAbsolute(self, relativeFilepath):

        dirname = os.path.dirname(__file__)
        return os.path.join(dirname, relativeFilepath)
    
    def removeQuestionsWithNoAnswer(self, Q_A_dict):
        Q_A_dictCopy = Q_A_dict.copy()

        for question, answer in Q_A_dict.items():
            if answer == '':
                Q_A_dictCopy.pop(question)
        return Q_A_dictCopy

    def removePersonalData(self, Q_A_dict):
        pass

    
