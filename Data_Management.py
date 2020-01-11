import os
import pyexcel
import pandas as pd
import pyexcel_ods3 as pods3

for file in os.listdir(os.getcwd()):
    if ".ods" in file:
        file1 = file[:-4]
        pyexcel.save_as(file_name = file, dest_file_name = file1 + ".csv")

z = []
for file in os.listdir(os.getcwd()):
    if ".csv" in file:
        z.append(file)

x = []
for i in range(0, len(z)):
    x.append(pd.read_csv(z[i]))

scores1 = []
scores2 = []
scores3 = []
scores4 = []
IELTS = []
Interview = []
for element in x:
    if 'Term 1' in list(element.columns.values):
        element = element[element['Term 1'].notnull()]
        global students
        students = element['Term 1'][1:element.shape[0]].tolist()
        for value in students:
            y = element.loc[element['Term 1'] == value]
            score = (float(y['Unnamed: 1'])*2 + float(y['Unnamed: 2'])*2 + float(y['Unnamed: 3']) + float(y['Unnamed: 4']) + float(y['Unnamed: 5']) + float(y['Unnamed: 6']) + float(y['Unnamed: 7']))/9
            scores1.append(score)
    if 'Term 2' in list(element.columns.values):
        element = element[element['Term 2'].notnull()]
        students = element['Term 2'][1:element.shape[0]].tolist()
        for value in students:
            y = element.loc[element['Term 2'] == value]
            score = (float(y['Unnamed: 1'])*2 + float(y['Unnamed: 2'])*2 + float(y['Unnamed: 3']) + float(y['Unnamed: 4']) + float(y['Unnamed: 5']) + float(y['Unnamed: 6']) + float(y['Unnamed: 7']))/9
            scores2.append(score)
    if 'Term 3' in list(element.columns.values):
        element = element[element['Term 3'].notnull()]
        students = element['Term 3'][1:element.shape[0]].tolist()
        for value in students:
            y = element.loc[element['Term 3'] == value]
            score = (float(y['Unnamed: 1'])*2 + float(y['Unnamed: 2'])*2 + float(y['Unnamed: 3']) + float(y['Unnamed: 4']) + float(y['Unnamed: 5']) + float(y['Unnamed: 6']) + float(y['Unnamed: 7']))/9
            scores3.append(score)
    if 'Term 4' in list(element.columns.values):
        element = element[element['Term 4'].notnull()]
        students = element['Term 4'][1:element.shape[0]].tolist()
        for value in students:
            y = element.loc[element['Term 4'] == value]
            score = (float(y['Unnamed: 1'])*2 + float(y['Unnamed: 2'])*2 + float(y['Unnamed: 3']) + float(y['Unnamed: 4']) + float(y['Unnamed: 5']) + float(y['Unnamed: 6']) + float(y['Unnamed: 7']))/9
            scores4.append(score)
    if 'IELTS' in list(element.columns.values):
        element = element[element['IELTS'].notnull()]
        students = element['IELTS'][1:element.shape[0]].tolist()
        for value in students:
            y = element.loc[element['IELTS'] == value]
            score = (float(y['Unnamed: 1']) + float(y['Unnamed: 2']) + float(y['Unnamed: 3']) + float(y['Unnamed: 4']))/4
            IELTS.append(score)
    if 'Interview' in list(element.columns.values):
        element = element[element['Interview'].notnull()]
        students = element['Interview'][1:element.shape[0]].tolist()
        for value in students:
            y = element.loc[element['Interview'] == value]
            score = (float(y['Unnamed: 1']) + float(y['Unnamed: 2']) + float(y['Unnamed: 3']) + float(y['Unnamed: 4']) + float(y['Unnamed: 5'])) / 5
            Interview.append(score)

final_scores = []
for value in students:
    x = []
    x.append(value)
    index1 = int(students.index(value))
    weighted_score = (((scores1[index1] + scores2[index1] + scores3[index1] + scores4[index1])/4)*0.4) + (((IELTS[index1]/9)*100)*0.3) + (((Interview[index1])/10)*100)*0.3
    x.append(str(weighted_score))
    final_scores.append(x)

final_scores.sort(key = lambda x: x[1], reverse = True)

with open('scores.txt', 'w') as f:
    for item in final_scores:
        f.write(str(item[0]) + ", " + str(item[1]) + '\n')


