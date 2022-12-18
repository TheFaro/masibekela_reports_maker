import pandas as pd
import warnings

from parseScoreSheets import parseScoreSheets
import class_lists
import calc_marks
from make_report import makeClassReports

warnings.simplefilter("ignore")


# workbook = xlrd.open_workbook('data/June/ScoreSheet2022-June-Form1A.xlsm')
# score_sheet = workbook.sheet_by_name("SCORE SHEET")
# print(score_sheet)

# df = pd.read_excel('data/June/ScoreSheet2022-June-Form1A.xlsm', sheet_name='SCORE SHEET', usecols='A:R', skiprows=5)
# df.drop(df.iloc[:, 0:3], inplace=True,axis=1)
# df.drop([0], inplace=True, axis=0)
# print(df.head())
# print(df)
# df.to_excel('form 1a.xlsx', sheet_name="June 2022")

# subject lists
secondary = []
high_school = []

# Function to open the classes excel files and calculates averages for each student in each subject


def readSubjects(path):
    df = pd.read_excel(path)
    # print(df.iloc[:,0:1])
    df.drop(df.iloc[:, 0:2], inplace=True, axis=1)

    return df.iloc[1].to_dict()


# parseScoreSheets()
class_lists.readStudents()
# print(class_lists.form1A)

# Term 1
# Form 1A averages

# secondary = readSubjects("data/parsed/Form1A.xlsx")
subjects = ['Name', 'English', 'LITinENG', 'SISWATI', 'MATH', 'INTER_SCIENCE', 'RELIGIOUS', 'ICT',
            'ADD_MATHS', 'GEOGRAPHY', 'HISTORY_', 'CONSUMER_SCIENCE', 'BUSINESS_STUDIES', 'BOOK_KEEPING', 'AGRICULTURE']
# calc_marks.calcTerm1Mark("data/parsed/Form1A.xlsx",
#                          class_lists.form1A, 'Form1A', subjects, secondary)
# calc_marks.calcTerm2Mark("data/parsed/Form1A.xlsx",
#                          class_lists.form1A, 'Form1A', subjects, secondary)
# calc_marks.calcCAMark("data/parsed/Form1A.xlsx",
#                       class_lists.form1A, 'Form1A', subjects, secondary)
# calc_marks.calcExamMark("data/parsed/Form1A.xlsx",
#                         class_lists.form1A, 'Form1A', subjects, secondary)
# calc_marks.calcFinalMark("data/average/Form1A.xlsx",
#                          class_lists.form1A, 'Form1A', subjects, secondary)
# calc_marks.calcRank("data/average/Form1A.xlsx",
#                     class_lists.form1A, 'Form1A', subjects, secondary)
# calc_marks.writeCategory("data/average/Form1A.xlsx",
#                          class_lists.form1A, 'Form1A', subjects, secondary)
# calc_marks.writeSubjectComments(
#     "data/average/Form1A.xlsx", class_lists.form1A, 'Form1A', subjects, secondary)
# makeClassReports('data/average/Form1A.xlsx', 'data/Attendance/Attendance_Conduct-2022-Term3-Form1A.xlsx',
#                  subjects, secondary, 'Form 1A', 50, 50, 55, 70, 'Mrs K.N. Dlamini', '17/01/2023', 'secondary')


# Form 1B averages
# secondary = readSubjects("data/parsed/Form1B.xlsx")
# calc_marks.calcTerm1Mark("data/parsed/Form1B.xlsx",
#                          class_lists.form1B, 'Form1B', subjects, secondary)
# calc_marks.calcTerm2Mark("data/parsed/Form1B.xlsx",
#                          class_lists.form1B, 'Form1B', subjects, secondary)
# calc_marks.calcCAMark("data/parsed/Form1B.xlsx",
#                       class_lists.form1B, 'Form1B', subjects, secondary)
# calc_marks.calcExamMark("data/parsed/Form1B.xlsx",
#                         class_lists.form1B, 'Form1B', subjects, secondary)
# calc_marks.calcFinalMark("data/average/Form1B.xlsx",
#                          class_lists.form1B, 'Form1B', subjects, secondary)
# calc_marks.calcRank("data/average/Form1B.xlsx",
#                     class_lists.form1B, 'Form1B', subjects, secondary)
# calc_marks.writeCategory("data/average/Form1B.xlsx",
#                          class_lists.form1B, 'Form1B', subjects, secondary)
# calc_marks.writeSubjectComments(
#     "data/average/Form1B.xlsx", class_lists.form1B, 'Form1B', subjects, secondary)
# makeClassReports('data/average/Form1B.xlsx', 'data/Attendance/Attendance_Conduct-2022-Term3-Form1B.xlsx',
#                  subjects, secondary, 'Form 1B', 50, 50, 55, 70, 'Ms B. Vilakati', '17/01/2023', 'secondary')

# Form 1C averages
secondary = readSubjects("data/parsed/Form1C.xlsx")
calc_marks.calcTerm1Mark("data/parsed/Form1C.xlsx",
                         class_lists.form1C, 'Form1C', subjects, secondary)
calc_marks.calcTerm2Mark("data/parsed/Form1C.xlsx",
                         class_lists.form1C, 'Form1C', subjects, secondary)
calc_marks.calcCAMark("data/parsed/Form1C.xlsx",
                      class_lists.form1C, 'Form1C', subjects, secondary)
calc_marks.calcExamMark("data/parsed/Form1C.xlsx",
                        class_lists.form1C, 'Form1C', subjects, secondary)
calc_marks.calcFinalMark("data/average/Form1C.xlsx",
                         class_lists.form1C, 'Form1C', subjects, secondary)
calc_marks.calcRank("data/average/Form1C.xlsx",
                    class_lists.form1C, 'Form1C', subjects, secondary)
calc_marks.writeCategory("data/average/Form1C.xlsx",
                         class_lists.form1C, 'Form1C', subjects, secondary)
calc_marks.writeSubjectComments(
    "data/average/Form1C.xlsx", class_lists.form1C, 'Form1C', subjects, secondary)
makeClassReports('data/average/Form1C.xlsx', 'data/Attendance/Attendance_Conduct-2022-Term3-Form1C.xlsx',
                 subjects, secondary, 'Form 1C', 50, 50, 55, 70, 'Mrs N. Ndzimandze', '17/01/2023', 'secondary')

# # # Form 2A averages
# secondary = readSubjects("data/parsed/Form2A.xlsx")
# calc_marks.calcTerm1Mark("data/parsed/Form2A.xlsx",
#                          class_lists.form2A, 'Form2A', subjects, secondary)
# calc_marks.calcTerm2Mark("data/parsed/Form2A.xlsx",
#                          class_lists.form2A, 'Form2A', subjects, secondary)
# calc_marks.calcCAMark("data/parsed/Form2A.xlsx",
#                       class_lists.form2A, 'Form2A', subjects, secondary)
# calc_marks.calcExamMark("data/parsed/Form2A.xlsx",
#                         class_lists.form2A, 'Form2A', subjects, secondary)
# calc_marks.calcFinalMark("data/average/Form2A.xlsx",
#                          class_lists.form2A, 'Form2A', subjects, secondary)
# calc_marks.calcRank("data/average/Form2A.xlsx",
#                     class_lists.form2A, 'Form2A', subjects, secondary)
# calc_marks.writeCategory("data/average/Form2A.xlsx",
#                          class_lists.form2A, 'Form2A', subjects, secondary)
# calc_marks.writeSubjectComments(
#     "data/average/Form2A.xlsx", class_lists.form2A, 'Form2A', subjects, secondary)
# makeClassReports('data/average/Form2A.xlsx', 'data/Attendance/Attendance_Conduct-2022-Term3-Form2A.xlsx',
#                  subjects, secondary, 'Form 2A', 50, 50, 55, 70, 'Ms T. Masuku', '17/01/2023', 'secondary')

# Form 2B averages
# secondary = readSubjects("data/parsed/Form2B.xlsx")
# calc_marks.calcTerm1Mark("data/parsed/Form2B.xlsx",
#                          class_lists.form2B, 'Form2B', subjects, secondary)
# calc_marks.calcTerm2Mark("data/parsed/Form2B.xlsx",
#                          class_lists.form2B, 'Form2B', subjects, secondary)
# calc_marks.calcCAMark("data/parsed/Form2B.xlsx",
#                       class_lists.form2B, 'Form2B', subjects, secondary)
# calc_marks.calcExamMark("data/parsed/Form2B.xlsx",
#                         class_lists.form2B, 'Form2B', subjects, secondary)
# calc_marks.calcFinalMark("data/average/Form2B.xlsx",
#                          class_lists.form2B, 'Form2B', subjects, secondary)
# calc_marks.calcRank("data/average/Form2B.xlsx",
#                     class_lists.form2B, 'Form2B', subjects, secondary)
# calc_marks.writeCategory("data/average/Form2B.xlsx",
#                          class_lists.form2B, 'Form2B', subjects, secondary)
# calc_marks.writeSubjectComments(
#     "data/average/Form2B.xlsx", class_lists.form2B, 'Form2B', subjects, secondary)
# makeClassReports("data/average/Form2B.xlsx", 'data/Attendance/Attendance_Conduct-2022-Term3-Form2B.xlsx', subjects, secondary, 'Form 2B', 50, 50, 55, 70, 'Mrs J. Tsabedze', '17/01/2023', 'secondary')

# Form 4A averages
# high_school = readSubjects("data/parsed/Form4A.xlsx")
# subjects = ["Name", "ENGLISH", "SISWATI", "MATH", "PHYSICAL_SCIENCE", "RELIGIOUS", "GEOGRAPHY",
#             "HISTORY_", "ICT", "FOODandNUTRITION", "BIOLOGY", "ACCOUNTS", "FF", "AGRICULTURE", "ECONOMICS"]
# calc_marks.calcTerm1Mark("data/parsed/Form4A.xlsx",
#                          class_lists.form4A, 'Form4A', subjects, high_school)
# calc_marks.calcTerm2Mark("data/parsed/Form4A.xlsx",
#                          class_lists.form4A, 'Form4A', subjects, high_school)
# calc_marks.calcCAMark("data/parsed/Form4A.xlsx",
#                       class_lists.form4A, 'Form4A', subjects, high_school)
# calc_marks.calcExamMark("data/parsed/Form4A.xlsx",
#                         class_lists.form4A, 'Form4A', subjects, high_school)
# calc_marks.calcFinalMark("data/average/Form4A.xlsx",
#                          class_lists.form4A, 'Form4A', subjects, high_school)
# calc_marks.calcRank("data/average/Form4A.xlsx",
#                     class_lists.form4A, 'Form4A', subjects, high_school)
# calc_marks.writeCategory("data/average/Form4A.xlsx",
#                          class_lists.form4A, 'Form4A', subjects, high_school)
# calc_marks.writeSubjectComments(
#     "data/average/Form4A.xlsx", class_lists.form4A, "Form4A", subjects, high_school)
# makeClassReports("data/average/Form4A.xlsx", "data/Attendance/Attendace_Conduct-2022-Term3-Form4A.xlsx",
#                  subjects, high_school, 'Form 4A', 60, 60, 60, 70, 'Ms H. Dube', '17/01/2023', 'high school')

# # Form 4B averages
# high_school = readSubjects("data/parsed/Form4B.xlsx")
# calc_marks.calcTerm1Mark("data/parsed/Form4B.xlsx", class_lists.form4B, 'Form4B', subjects, high_school)
# calc_marks.calcTerm2Mark("data/parsed/Form4B.xlsx", class_lists.form4B, 'Form4B', subjects, high_school)
# calc_marks.calcCAMark("data/parsed/Form4B.xlsx", class_lists.form4B, 'Form4B', subjects, high_school)
# calc_marks.calcExamMark("data/parsed/Form4B.xlsx", class_lists.form4B, 'Form4B', subjects, high_school)
# calc_marks.calcFinalMark("data/average/Form4B.xlsx", class_lists.form4B, 'Form4B', subjects, high_school)
# calc_marks.calcRank("data/average/Form4B.xlsx", class_lists.form4B, 'Form4B', subjects, high_school)
# calc_marks.writeCategory("data/average/Form4B.xlsx", class_lists.form4B, 'Form4B', subjects, high_school)
# calc_marks.writeSubjectComments("data/average/Form4B.xlsx", class_lists.form4B, "Form4B", subjects, high_school)
# makeClassReports("data/average/Form4B.xlsx", "data/Attendance/Attendance_Conduct-2022-Term3-Form4B.xlsx", subjects, high_school, 'Form 4B', 60, 60, 60, 70, 'Ms L.P. Dlamini', '17/01/2023', 'high school)
