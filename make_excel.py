import pandas as pd

excel_filename = "academic_data.xlsx"

print("Starting to build the Excel file...")

subjects_data = {
    'Code': ['IF101', 'IF102', 'IF103'],
    'Subject Name': ['Programming Fundamentals', 'Database Systems', 'Web Programming'],
    'SKS': [3, 3, 2]
}
df_subjects = pd.DataFrame(subjects_data)

students_data = {
    'StudentID': ['S001', 'S002', 'S003'],
    'Student Name': ['Ahmad', 'Budi', 'Cindy'],
    'Group': ['A', 'A', 'B']
}
df_students = pd.DataFrame(students_data)

scores_data = {
    'ID': [1, 2, 3, 4],
    'SubjectCode': ['IF101', 'IF102', 'IF101', 'IF103'],
    'StudentID': ['S001', 'S001', 'S002', 'S003'],
    'Score': [85, 78, 90, 88]
}
df_scores = pd.DataFrame(scores_data)

with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
    df_subjects.to_excel(writer, sheet_name='Subjects', index=False)
    df_students.to_excel(writer, sheet_name='Students', index=False)
    df_scores.to_excel(writer, sheet_name='RawScores', index=False)

print(f"🎉 Success! '{excel_filename}' has been created with all 3 sheets matching the professor's design.")