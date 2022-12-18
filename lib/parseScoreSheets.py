from openpyxl import load_workbook
import pandas as pd


def parseScoreSheets(): 
    classes = ["Form1A", "Form1B", "Form1C", "Form2A", "Form2B", "Form4A", "Form4B"]
    months = ['May', 'June', 'Mock', 'October','November', 'Final']
    
    for month in months:
        for stream in classes:
            # print("data/%s/ScoreSheet2022-%s-%s.xlsm" % (month, month, stream))
            path = "data/%s/ScoreSheet2022-%s-%s.xlsm" % (month, month, stream)
            # print(path)
            
            # Read data from the score sheet
            df = pd.read_excel(path, sheet_name='SCORE SHEET', usecols='A:R', skiprows=5)
            
            # drop uneeded columns 
            df.drop(df.iloc[:, 0:3], inplace=True,axis=1)
            
            try: 
                parsedPath = "data/parsed/%s.xlsx" % (stream)
                book = load_workbook(parsedPath)
                writer = pd.ExcelWriter(parsedPath, engine='openpyxl')
                writer.book = book
                
                
                df.to_excel(writer, sheet_name="%s" % (month) )
                writer.close()
            except FileNotFoundError as e:
                print(e)
                # write data in excel file in month sheet
                df.to_excel("data/parsed/%s.xlsx" % (stream), sheet_name="%s" % (month), )
