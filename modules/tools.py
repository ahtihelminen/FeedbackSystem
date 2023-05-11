
import csv
import json
import pandas as pd
import os
from docx import Document
from docx.shared import Inches


class Tools:
    def __init__(self) -> None:
        pass

    def convertXlsxToCsv(self, excelFeedackFile):
        excelFilepathRel = f"../feedbacksExcel/{excelFeedackFile}"
        excelFilepathAbs=self.relativeFilepathToAbsolute(excelFilepathRel)

        dataFrameToBeConverted = pd.read_excel(excelFilepathAbs)
        
        csvFilepathRel = f"../feedbacksCsv/{excelFeedackFile.replace('.xlsx', '.csv')}"
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

    def createWordDoc(self, rows, cols):
        
        # create a new document
        document = Document()

        # add a table
        table = document.add_table(rows=3, cols=3)

        # populate the table
        table.cell(0, 0).text = 'Name'
        table.cell(0, 1).text = 'Age'
        table.cell(0, 2).text = 'Gender'

        table.cell(1, 0).text = 'Alice'
        table.cell(1, 1).text = '25'
        table.cell(1, 2).text = 'Female'

        table.cell(2, 0).text = 'Bob'
        table.cell(2, 1).text = '30'
        table.cell(2, 2).text = 'Male'

        # adjust the column widths
        for row in table.rows:
            for cell in row.cells:
                cell.width = Inches(2.0)

        # save the document
        document.save('example.docx')

