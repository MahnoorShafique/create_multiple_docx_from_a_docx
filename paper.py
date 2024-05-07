import csv
import os
from docx import Document
from docx.shared import Pt
# Function to write content to a new doc file
def write_docx_file(file_path, document):
    document.save(file_path)

english_exam_file_path = "Checkpoint- English Grade 1.docx"


csv_file_path = "student_data_niete.csv"


output_directory = "english_paper_dir"


if not os.path.exists(output_directory):
    os.makedirs(output_directory)


students_by_emis = {}
with open(csv_file_path, 'r', encoding='utf-8') as csv_file:
    csv_reader = csv.DictReader(csv_file)
    for row in csv_reader:
        emis = row['emis']
        if emis not in students_by_emis:
            students_by_emis[emis] = []
        students_by_emis[emis].append(row)


for emis, students in students_by_emis.items():
    emis_folder = os.path.join(output_directory, f"EMIS_{emis}")
    if not os.path.exists(emis_folder):
        os.makedirs(emis_folder)

    for student in students:
        
        student_name = student['student_name']
        admission_number = student['admission_number']
        
        
        output_file_name = f"english_{admission_number}.docx"
        output_file_path = os.path.join(emis_folder, output_file_name)
        
       
        doc = Document(english_exam_file_path)
        
        
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if 'student_name' in run.text:
                    run.text = run.text.replace('student_name', student_name)
                    run.font.size = Pt(15)
                    run.font.name = 'Arial'
                if 'admission_number' in run.text:
                    run.text = run.text.replace('admission_number', admission_number)
                    run.font.size = Pt(15)
                    run.font.name = 'Arial'
                if 'emis' in run.text:
                    run.text = run.text.replace('emis', emis)
                    run.font.size = Pt(15)
                    run.font.name = 'Arial'
        
       
        write_docx_file(output_file_path, doc)

print("Doc files generated successfully.")
