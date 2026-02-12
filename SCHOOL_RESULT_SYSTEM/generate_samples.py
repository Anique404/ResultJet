import sys
import os
sys.path.insert(0, '.')
import pandas as pd
from ResultSystem import calculate_student_result, detect_subjects_and_max_marks, generate_result_card_pdf

# Generate sample PDFs from SAMPLE_FILES
print('Generating sample PDFs from SAMPLE_FILES...')
os.makedirs('SAMPLE_OUTPUT', exist_ok=True)

for class_name, sample_file in [('KG', 'KG_sample.xlsx'), ('5th', '5th_sample.xlsx'), ('10th', '10th_sample.xlsx')]:
    excel_file = os.path.join('SAMPLE_FILES', sample_file)
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        subjects = detect_subjects_and_max_marks(df)
        
        output_subdir = os.path.join('SAMPLE_OUTPUT', f'{class_name}_results')
        os.makedirs(output_subdir, exist_ok=True)
        
        generated = 0
        for idx, student_row in enumerate(df.to_dict('records')):
            try:
                student_result = calculate_student_result(pd.Series(student_row), subjects)
                pdf_filename = f"{student_result['roll_no']:03d}_{student_result['name'].replace(' ', '_')}_{class_name}_Result.pdf"
                pdf_path = os.path.join(output_subdir, pdf_filename)
                if generate_result_card_pdf(student_result, class_name, pdf_path):
                    generated += 1
            except Exception as e:
                print(f'  Error: {e}')
        
        print(f'✅ {class_name}: Generated {generated} sample PDFs in SAMPLE_OUTPUT/{class_name}_results/')
    else:
        print(f'⚠️ Sample file not found: {excel_file}')

print('\n✅ Sample PDFs created successfully!')
