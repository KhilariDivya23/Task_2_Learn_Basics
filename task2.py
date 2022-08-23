import pandas as pd


class TestDetails:
    def __init__(self, name, score, time_taken, answered, correct, wrong, skipped):
        self.name = name
        self.score = score
        self.time_taken = time_taken
        self.answered = answered
        self.correct = correct
        self.wrong = wrong
        self.skipped = skipped


class MainDetails:
    def __init__(self, student_name, student_id, chapter_tag):
        self.student_name = student_name
        self.student_id = student_id
        self.chapter_tag = chapter_tag
        self.tests = []


sheets = pd.read_excel('Input_2 - Python Developer Intern - Task 2 - Datasets.xlsx')
df = pd.DataFrame(sheets)

students = []
for i in df.index:
    student = MainDetails(df['Name'][i], df['id'][i], df['Chapter Tag'][i])
    for j in range(3, len(df.columns), 6):
        if df[df.columns[j]][i] == '-' or df[df.columns[j+1]][i] == '-' or df[df.columns[j+2]][i] == '-' or df[df.columns[j+3]][i] == '-' or df[df.columns[j+4]][i] == '-' or df[df.columns[j+5]][i] == '-':
            continue
        test = TestDetails(df.columns[j].split('-')[0],
                           df[df.columns[j]][i],
                           df[df.columns[j+1]][i],
                           df[df.columns[j+2]][i],
                           df[df.columns[j+3]][i],
                           df[df.columns[j+4]][i],
                           df[df.columns[j+5]][i])
        student.tests.append(test)
    students.append(student)

# print(len(students[2].tests))

write = pd.ExcelWriter('output2.xlsx', engine='xlsxwriter')
data = {'Name': [],
        'Username': [],
        'Chapter Tag': [],
        'Test Name': [],
        'answered': [],
        'correct': [],
        'score': [],
        'skipped': [],
        'time_taken': [],
        'wrong': []}

for i in students:
    for j in i.tests:
        data['Name'].append(i.student_name)
        data['Username'].append(i.student_id)
        data['Chapter Tag'].append(i.chapter_tag)
        data['Test Name'].append(j.name)
        data['answered'].append(j.answered)
        data['correct'].append(j.correct)
        data['score'].append(j.score)
        data['skipped'].append(j.skipped)
        data['time_taken'].append(j.time_taken)
        data['wrong'].append(j.wrong)

new_df = pd.DataFrame(data)
new_df.to_excel(write, sheet_name='in', index=False)
write.save()














