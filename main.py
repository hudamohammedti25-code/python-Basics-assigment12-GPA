import pandas as pd
import numpy as np

excel_file = "academic_data.xlsx"

print("==================================================")
print("🎓 Nusa Putra University - Academic Data Management")
print("==================================================")

try:
    print("\n[Step 1] Reading data from workbook sheets...")
    df_subjects = pd.read_excel(excel_file, sheet_name='Subjects')
    df_students = pd.read_excel(excel_file, sheet_name='Students')
    df_scores = pd.read_excel(excel_file, sheet_name='RawScores')

    df_subjects.columns = df_subjects.columns.str.strip()
    df_students.columns = df_students.columns.str.strip()
    df_scores.columns = df_scores.columns.str.strip()

    print("\n--- Subjects Master Data ---")
    print(df_subjects.to_string(index=False))

    print("\n--- Students Master Data ---")
    print(df_students.to_string(index=False))

    print("\n[Step 3] Merging sheets via StudentID and SubjectCode...")
    
    df_scores['StudentID'] = df_scores['StudentID'].astype(str).str.strip()
    df_students['StudentID'] = df_students['StudentID'].astype(str).str.strip()
    df_scores['SubjectCode'] = df_scores['SubjectCode'].astype(str).str.strip()
    df_subjects['Code'] = df_subjects['Code'].astype(str).str.strip()

    merged_students = pd.merge(df_scores, df_students, on='StudentID', how='inner')
    
    final_merged_df = pd.merge(merged_students, df_subjects, left_on='SubjectCode', right_on='Code', how='inner')

    if final_merged_df.empty:
        print("\n⚠️ Warning: Merged dataframe is empty! Please check if StudentIDs match exactly between sheets.")
        exit()

    def calculate_grade_and_points(score):
        if score >= 80:
            return 'A', 4.0
        elif score >= 75:
            return 'B+', 3.5
        elif score >= 70:
            return 'B', 3.0
        elif score >= 65:
            return 'C+', 2.5
        elif score >= 60:
            return 'C', 2.0
        elif score >= 50:
            return 'D', 1.0
        else:
            return 'E', 0.0

    print("\n[Step 4] Converting numeric scores to Grades and Quality Points...")
    grade_results = final_merged_df['Score'].apply(calculate_grade_and_points)
    final_merged_df['Grade'] = [res[0] for res in grade_results]
    final_merged_df['GradePoints'] = [res[1] for res in grade_results]

    final_merged_df['WeightedPoints'] = final_merged_df['GradePoints'] * final_merged_df['SKS']

    print("\n[Step 5] Calculating total SKS and weighted GPA (IPK)...")
    
    student_name_col = 'Student Name' if 'Student Name' in final_merged_df.columns else 'StudentName'

    student_summary = final_merged_df.groupby(['StudentID', student_name_col]).agg(
        Total_SKS=('SKS', 'sum'),
        Total_Weighted_Points=('WeightedPoints', 'sum')
    ).reset_index()

    student_summary['GPA_IPK'] = round(student_summary['Total_Weighted_Points'] / student_summary['Total_SKS'], 2)

    gpa_report = student_summary[['StudentID', student_name_col, 'Total_SKS', 'GPA_IPK']]
    gpa_report.columns = ['StudentID', 'Student Name', 'Total SKS', 'GPA / IPK']

    print("\n--- Final Students GPA (IPK) Report ---")
    print(gpa_report.to_string(index=False))

    class_avg = gpa_report['GPA / IPK'].mean()
    print("\n==================================================")
    print(f"📊 Overall Class Average GPA (IPK): {class_avg:.2f}")
    print("==================================================")

    print(f"\n[Step 8] Exporting the final report back to Excel...")
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        gpa_report.to_excel(writer, sheet_name='GPA_Report', index=False)

    print("✅ Data successfully saved in sheet 'GPA_Report' within 'academic_data.xlsx'!")
    print("==================================================")

except Exception as e:
    print(f"\n❌ An error occurred: {e}")
    print("Tip: Make sure to run 'make_excel.py' first to reset the data format correctly.")