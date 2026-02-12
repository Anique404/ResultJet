import sys
import os
sys.path.insert(0, '.')
import pandas as pd
from ResultSystem import detect_subjects_and_max_marks, calculate_student_result, generate_result_card_pdf

excel_file = 'SAMPLE_FILES/KG_sample.xlsx'
df = pd.read_excel(excel_file)
print('Loaded Excel file:', excel_file)
print('Shape:', df.shape)

subjects = detect_subjects_and_max_marks(df)
print('Detected subjects:', subjects)

# Try to calculate result for first student
student_row = df.iloc[0]
student_result = calculate_student_result(student_row, subjects)
print('\nStudent result:')
print(f'  Name: {student_result["name"]}')
print(f'  Overall Status: {student_result["overall_status"]}')
print(f'  Subjects: {len(student_result["subjects"])}')
print(f'  Overall Percentage: {student_result["overall_percentage"]}')

# Try to generate PDF
os.makedirs('SAMPLE_OUTPUT/KG_results', exist_ok=True)
pdf_path = 'SAMPLE_OUTPUT/KG_results/001_Ahmed_Ali_KG_Result.pdf'
success = generate_result_card_pdf(student_result, 'KG', pdf_path)
print(f'\nPDF Generation: {"Success" if success else "Failed"}')
print(f'PDF Path: {pdf_path}')

if os.path.exists(pdf_path):
    file_size = os.path.getsize(pdf_path)
    print(f'PDF Size: {file_size} bytes')
