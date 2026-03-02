"""
The Blessing School - RESULT CARD GENERATION SYSTEM v3.0
=========================================================
Premium Version with Position Ranking & Validation
Author: School Administration System
"""

import sys
import os
import pandas as pd
from datetime import datetime
import threading
from typing import List, Dict, Tuple, Optional
import tkinter as tk
from tkinter import ttk, messagebox
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import subprocess

# ============================================================================
# CONFIGURATION
# ============================================================================

PASS_THRESHOLD = 33
SCHOOL_NAME = "The Blessing School"
SCHOOL_CITY = "Karachi, Pakistan"

CLASSES = [
    "Playgroup (RED)", "Playgroup (BLUE)",
    "KG-1 (RED)", "KG-1 (BLUE)",
    "KG-2 (RED)", "KG-2 (BLUE)",
    "ONE", "TWO", "THREE", "FOUR", "FIVE",
    "SIXTH", "SEVEN", "EIGHT", "NINE", "TEN"
]

# School colors
PRIMARY_COLOR = colors.HexColor('#1a3b5d')  # Navy Blue
SECONDARY_COLOR = colors.HexColor('#c49a1c')  # Gold
LIGHT_BG = colors.HexColor('#f8f9fa')
SUCCESS_COLOR = colors.HexColor('#28a745')
DANGER_COLOR = colors.HexColor('#dc3545')

# ============================================================================
# CORE FUNCTIONS
# ============================================================================

def detect_subjects_and_max_marks(df: pd.DataFrame) -> List[Tuple[str, str, Optional[str]]]:
    """Detects subjects and their max marks from Excel columns."""
    fixed_columns = set(['RollNo', 'StudentName'])
    
    subjects = []
    for col in df.columns:
        if col in fixed_columns or col.endswith('_Max'):
            continue
        
        max_col = col + '_Max'
        if max_col in df.columns:
            subjects.append((col, col, max_col))
        else:
            subjects.append((col, col, None))
    
    return subjects


def calculate_subject_percentage(obtained: float, max_marks: float) -> float:
    """Calculate percentage: (obtained / max) * 100"""
    if max_marks == 0:
        return 0.0
    return round((obtained / max_marks) * 100, 2)


def is_subject_pass(percentage: float) -> bool:
    """Return True if percentage >= PASS_THRESHOLD"""
    return percentage >= PASS_THRESHOLD


def get_grade(percentage: float) -> str:
    """Assign grade based on percentage."""
    if percentage >= 90:
        return "A+"
    elif percentage >= 80:
        return "A"
    elif percentage >= 70:
        return "B"
    elif percentage >= 60:
        return "C"
    elif percentage >= 50:
        return "D"
    elif percentage >= PASS_THRESHOLD:
        return "E"
    else:
        return "F"


def determine_overall_result(subject_percentages: List[float]) -> bool:
    """PASS only if ALL subjects are >= PASS_THRESHOLD"""
    return all(p >= PASS_THRESHOLD for p in subject_percentages)


def calculate_student_result(student_row: pd.Series, subjects: List[Tuple[str, str, Optional[str]]]) -> Optional[Dict]:
    """
    Calculate complete result for a student.
    Returns None if data is invalid.
    """
    try:
        # Validate required fields
        if pd.isna(student_row.get('RollNo')) or pd.isna(student_row.get('StudentName')):
            return None
        
        roll_no = int(student_row['RollNo'])
        name = str(student_row['StudentName']).strip()
        
        if not name:
            return None
        
        result = {
            'roll_no': roll_no,
            'name': name,
            'subjects': [],
            'total_obtained': 0,
            'total_max': 0,
            'failed_subjects': [],
            'is_valid': True
        }
        
        subject_percentages = []
        
        for subject_name, marks_col, max_col in subjects:
            # Check if marks exist and are valid
            if pd.isna(student_row.get(marks_col)):
                return None  # Missing marks = invalid student
            
            try:
                obtained = float(student_row[marks_col])
                
                # Get max marks
                if max_col and not pd.isna(student_row.get(max_col)):
                    max_marks = float(student_row[max_col])
                else:
                    max_marks = 100.0
                
                # Validate marks
                if obtained < 0 or obtained > max_marks:
                    return None
                
                percentage = calculate_subject_percentage(obtained, max_marks)
                subject_percentages.append(percentage)
                
                is_pass = is_subject_pass(percentage)
                status = "PASS" if is_pass else "FAIL"
                
                if not is_pass:
                    result['failed_subjects'].append(f"{subject_name}")
                
                grade = get_grade(percentage)
                
                result['subjects'].append({
                    'name': subject_name,
                    'obtained': obtained,
                    'max': max_marks,
                    'percentage': percentage,
                    'grade': grade,
                    'status': status
                })
                
                result['total_obtained'] += obtained
                result['total_max'] += max_marks
                
            except (ValueError, TypeError):
                return None
        
        # Calculate overall results
        if result['total_max'] > 0:
            result['overall_percentage'] = calculate_subject_percentage(result['total_obtained'], result['total_max'])
            result['overall_grade'] = get_grade(result['overall_percentage'])
            result['overall_status'] = "PASS" if determine_overall_result(subject_percentages) else "FAIL"
        else:
            return None
        
        return result
        
    except Exception:
        return None


def calculate_positions(students_results: List[Dict]) -> List[Dict]:
    """Calculate position/rank for each student based on percentage."""
    if not students_results:
        return students_results
    
    # Sort by percentage (descending)
    sorted_students = sorted(students_results, key=lambda x: x['overall_percentage'], reverse=True)
    
    # Assign positions
    position = 1
    prev_percentage = None
    
    for i, student in enumerate(sorted_students):
        if prev_percentage is not None and student['overall_percentage'] == prev_percentage:
            student['position'] = position  # Same position for same percentage
        else:
            position = i + 1
            student['position'] = position
        
        prev_percentage = student['overall_percentage']
    
    # Add suffix to position (1st, 2nd, 3rd, 4th, etc.)
    for student in students_results:
        pos = student['position']
        if pos == 1:
            student['position_str'] = "1st"
        elif pos == 2:
            student['position_str'] = "2nd"
        elif pos == 3:
            student['position_str'] = "3rd"
        else:
            student['position_str'] = f"{pos}th"
    
    return students_results


# ============================================================================
# PDF GENERATION - PREMIUM DESIGN
# ============================================================================

def create_result_card(student_result: Dict, class_name: str) -> Table:
    """
    Create a single result card with premium design.
    """
    card_elements = []
    
    # ===== LOGO CENTER MEIN =====
    card_elements.append(Spacer(1, 0.1*inch))

    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")

    if os.path.exists(logo_path):
        # Logo center mein lagao
        logo = Image(logo_path, width=1.1*inch, height=1.1*inch)
        
        # Logo ko center karne ke liye table use karo
        logo_table = Table([[logo]], colWidths=[7.2*inch])
        logo_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ]))
        logo_table.hAlign = 'CENTER'
        card_elements.append(logo_table)
        
        # Logo ke neeche school name
        school_style = ParagraphStyle(
            'SchoolName',
            fontSize=17,
            fontName='Times-Bold',
            textColor=PRIMARY_COLOR,
            alignment=TA_CENTER,
            spaceAfter=4
        )
        school_name = Paragraph(SCHOOL_NAME, school_style)
        card_elements.append(school_name)
        card_elements.append(Spacer(1, 0.2*inch))  # ← YEH LINE ADD KARO - space dene ke liye

        
    else:
        # Agar logo nahi hai to sirf school name center mein
        school_style = ParagraphStyle(
            'SchoolName',
            fontSize=15,
            fontName='Times-Bold',
            textColor=PRIMARY_COLOR,
            alignment=TA_CENTER,
        )
        card_elements.append(Paragraph(SCHOOL_NAME, school_style))
        

    
    # ===== SCHOOL ADDRESS =====
    address_style = ParagraphStyle(
        'Address',
        fontSize=8,
        fontName='Helvetica',
        alignment=TA_CENTER,
        textColor=colors.grey,
        spaceAfter=4
    )
    address_text = "Main Road Alfalah Town Sadik Abad Ph #0300-6721263"
    card_elements.append(Paragraph(address_text, address_style))
    
    # ===== MID TERM EXAMINATION TITLE =====
    exam_title_style = ParagraphStyle(
        'ExamTitle',
        fontSize=10,
        fontName='Times-Bold',
        alignment=TA_CENTER,
        textColor=PRIMARY_COLOR,
        spaceAfter=8
    )
    card_elements.append(Paragraph("First / Mid / Final Term Examination 2026", exam_title_style))
    
    # ===== NAME (LEFT) AND CLASS (RIGHT) IN SAME LINE =====
    name_class_data = []
    
    # Name (left)
    name_para = Paragraph(
        f"<b>Name:</b> {student_result['name']}",
        ParagraphStyle('Name', fontSize=10, fontName='Helvetica', alignment=TA_LEFT)
    )
    
    # Class (right)
    class_para = Paragraph(
        f"<b>Class:</b> {class_name}",
        ParagraphStyle('Class', fontSize=10, fontName='Times-Roman', alignment=TA_RIGHT)
    )
    
    name_class_data.append([name_para, class_para])
    
    name_class_table = Table(name_class_data, colWidths=[3.6*inch, 3.6*inch])
    name_class_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('RIGHTPADDING', (1, 0), (1, 0), 55),  # Class ko right se padding
    ]))
    name_class_table.hAlign = 'CENTER'
    card_elements.append(name_class_table)
    
    # ===== DATE (RIGHT) =====
    date_para = Paragraph(
        f"<b>Date:</b> {datetime.now().strftime('%d/%m/%Y')}",
        ParagraphStyle('Date', fontSize=10, fontName='Times-Roman', alignment=TA_RIGHT)
    )
    
    date_table = Table([[date_para]], colWidths=[7.2*inch])
    date_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, 0), 'RIGHT'),
        ('RIGHTPADDING', (0, 0), (0, 0), 40),  # Date ko right se padding
    ]))
    date_table.hAlign = 'CENTER'
    card_elements.append(date_table)
    
    card_elements.append(Spacer(1, 0.1*inch))
    
    # ===== SUBJECTS TABLE WITH TOTAL =====
    subject_data = [['Subject', 'Marks', 'Max', 'Percentage', 'Grade', 'Status']]
    
    for sub in student_result['subjects']:
        subject_data.append([
            sub['name'],
            f"{sub['obtained']:.0f}",
            f"{sub['max']:.0f}",
            f"{sub['percentage']:.0f}%",
            sub['grade'],
            sub['status']
        ])
    
    # TOTAL ROW ADD KARO
    subject_data.append([
        'TOTAL',
        f"{student_result['total_obtained']:.0f}",
        f"{student_result['total_max']:.0f}",
        f"{student_result['overall_percentage']:.0f}%",
        student_result['overall_grade'],
        ''
    ])
    
    subject_table = Table(subject_data, colWidths=[1.5*inch, 1.0*inch, 1.0*inch, 1.0*inch, 1.0*inch, 1.5*inch])
    
    subject_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, 0), 'Times-Bold', 9),
        ('FONT', (0, 1), (-1, -2), 'Times-Roman', 8),        # Data font (excluding last row)
        ('FONT', (0, -1), (-1, -1), 'Times-Bold', 9),  # Total row bold
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY_COLOR),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#000000")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.white]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.white),        # Total row background
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    
    # # Color coding for status
    # for i, sub in enumerate(student_result['subjects'], start=1):
    #     if sub['status'] == 'PASS':
    #         subject_table.setStyle(TableStyle([('TEXTCOLOR', (5, i), (5, i), SUCCESS_COLOR)]))
    #     else:
    #         subject_table.setStyle(TableStyle([('TEXTCOLOR', (5, i), (5, i), DANGER_COLOR)]))
    
    subject_table.hAlign = 'CENTER'
    card_elements.append(subject_table)
    card_elements.append(Spacer(1, 0.1*inch))
    
    # ===== PERCENTAGE (LEFT) AND GRADE (RIGHT) IN SAME LINE =====
    summary_data = []
    
    # Percentage (left)
    percentage_style = ParagraphStyle(
        'Percentage',
        fontSize=9,
        fontName='Times-Bold',
        alignment=TA_LEFT,
    )
    percentage_text = f"<b>Percentage: {student_result['overall_percentage']:.1f}%</b>"
    percentage_para = Paragraph(percentage_text, percentage_style)
    
    # Grade (right)
    grade_style = ParagraphStyle(
        'Grade',
        fontSize=9,
        fontName='Times-Bold',
        alignment=TA_RIGHT,
    )
    grade_text = f"<b>Grade: {student_result['overall_grade']}</b>"
    grade_para = Paragraph(grade_text, grade_style)
    
    # Dono ek hi row mein
    summary_data.append([percentage_para, grade_para])
    
    summary_table = Table(summary_data, colWidths=[3.6*inch, 3.6*inch])
    summary_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('LEFTPADDING', (0, 0), (0, 0), 10),
        ('RIGHTPADDING', (1, 0), (1, 0), 40),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
    ]))
    summary_table.hAlign = 'CENTER'
    card_elements.append(summary_table)
    card_elements.append(Spacer(1, 0.1*inch))
    
    # ===== POSITION (LEFT) AND COMMENTS (RIGHT) IN SAME LINE =====
    pr_data = []
    
    # Position (left)
    position_color = SECONDARY_COLOR if student_result['position'] <= 3 else PRIMARY_COLOR
    position_text = f"<b>POSITION: {student_result['position_str']}</b>"
    position_para = Paragraph(position_text, ParagraphStyle(
        'Position',
        fontSize=9,
        fontName='Times-Bold',
        textColor=position_color,
        alignment=TA_LEFT
    ))
    
    # Comments (right)
    comments_style = ParagraphStyle(
        'Comments',
        fontSize=9,
        fontName='Times-Roman',
        alignment=TA_RIGHT,
        textColor=colors.black
    )
    comments_text = "Comments: ____________________________"
    comments_para = Paragraph(comments_text, comments_style)
    
    # Dono ek hi row mein
    pr_data.append([position_para, comments_para])
    
    pr_table = Table(pr_data, colWidths=[3.5*inch, 3.5*inch])
    pr_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('LEFTPADDING', (0, 0), (0, 0), 10),
        ('RIGHTPADDING', (1, 0), (1, 0), 20),
    ]))
    pr_table.hAlign = 'CENTER'
    card_elements.append(pr_table)
    
    # ===== TEACHER AND PRINCIPAL SIGNATURE IN SAME LINE =====
    card_elements.append(Spacer(1, 0.1*inch))  # Medium space (1/5 inch)

    signature_data = []
    
    # Teacher Signature
    teacher_style = ParagraphStyle(
        'TeacherSignature',
        fontSize=10,
        fontName='Times-Roman',
        alignment=TA_LEFT,
        textColor=colors.grey
    )
    teacher_text = "Teacher Signature: ____________________"
    teacher_para = Paragraph(teacher_text, teacher_style)
    
    # Principal Signature
    principal_style = ParagraphStyle(
        'PrincipalSignature',
        fontSize=9,
        fontName='Times-Roman',
        alignment=TA_RIGHT,
        textColor=colors.grey
    )
    principal_text = "Principal Signature: ____________________"
    principal_para = Paragraph(principal_text, principal_style)
    
    # Dono signatures ek hi row mein
    signature_data.append([teacher_para, principal_para])
    
    signature_table = Table(signature_data, colWidths=[3.6*inch, 3.6*inch])
    signature_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('LEFTPADDING', (0, 0), (0, 0), 10),
        ('RIGHTPADDING', (1, 0), (1, 0), 40),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
    ]))
    signature_table.hAlign = 'CENTER'
    card_elements.append(signature_table)
    card_elements.append(Spacer(1, 0.1*inch))
    
    # ===== WRAP EVERYTHING IN A BORDER =====
    card_table = Table([[e] for e in card_elements])
    card_table.setStyle(TableStyle([
        # ('BOX', (0, 0), (-1, -1), 1, PRIMARY_COLOR),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    
    card_table.hAlign = 'CENTER'
    return card_table

def generate_result_pdf(students_results: List[Dict], class_name: str, output_path: str) -> bool:
    """
    Generate PDF with 2 cards per page (A4 Portrait) - Cards separated by line
    """
    try:
        doc = SimpleDocTemplate(output_path, 
                               pagesize=A4,
                               rightMargin=0.4*inch, 
                               leftMargin=0.4*inch,
                               topMargin=0.3*inch, 
                               bottomMargin=0.3*inch)
        
        elements = []
        
        # Calculate height for each card to fit 2 per page
        page_height = A4[1]  # 841.89 points
        available_height = page_height - 0.6*inch  # Remove margins
        card_height = available_height / 2  # Divide by 2 for 2 cards
        
        for i in range(0, len(students_results), 2):
            # Card 1 - Top half
            card1 = create_result_card(students_results[i], class_name)
            elements.append(card1)
            
            # Check if there's a second card
            if i + 1 < len(students_results):
                # Add space before line
                elements.append(Spacer(1, 0.2*inch))
                
                # ===== LINE BETWEEN CARDS =====
                line_style = ParagraphStyle(
                    'SeparatorLine',
                    fontSize=1,
                    alignment=TA_CENTER,
                    textColor=colors.darkblue,
                )
                # Create a line that spans the page width
                line_text = "_" * 900  # 100 underscores
                line_para = Paragraph(line_text, line_style)
                
                # Center the line
                line_table = Table([[line_para]], colWidths=[7.2*inch])
                line_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                ]))
                elements.append(line_table)
                
                # Add space after line
                elements.append(Spacer(1, 0.2*inch))
                
                # Card 2 - Bottom half
                card2 = create_result_card(students_results[i+1], class_name)
                elements.append(card2)
            
            # Page break after every 2 cards
            if i + 2 < len(students_results):
                elements.append(PageBreak())
        
        doc.build(elements)
        return True
        
    except Exception as e:
        print(f"PDF Error: {e}")
        return False


# ============================================================================
# EXCEL PROCESSING
# ============================================================================

def load_excel_file(file_path: str) -> Optional[Tuple[pd.DataFrame, List[Tuple[str, str, Optional[str]]]]]:
    """Load and validate Excel file."""
    try:
        df = pd.read_excel(file_path)
        
        required_cols = ['RollNo', 'StudentName']
        for col in required_cols:
            if col not in df.columns:
                return None
        
        subjects = detect_subjects_and_max_marks(df)
        
        if not subjects:
            return None
        
        return df, subjects
        
    except Exception:
        return None


def process_class_to_pdf(class_name: str, class_data_dir: str, output_dir: str, 
                         single_roll_no: Optional[int] = None, 
                         progress_callback=None) -> Tuple[int, int, List[str]]:
    """Process class and generate PDF with valid students only."""
    
    excel_file = os.path.join(class_data_dir, f"{class_name}.xlsx")
    
    if not os.path.exists(excel_file):
        return 0, 0, [f"File not found: {class_name}.xlsx"]
    
    result = load_excel_file(excel_file)
    if result is None:
        return 0, 0, ["Invalid Excel format"]
    
    df, subjects = result
    
    output_subdir = os.path.join(output_dir, f"{class_name}_results")
    os.makedirs(output_subdir, exist_ok=True)
    
    # Process all students
    students = df.to_dict('records')
    valid_results = []
    invalid_count = 0
    errors = []
    
    for idx, student_row in enumerate(students):
        try:
            # Filter by roll number if specified
            if single_roll_no and int(student_row['RollNo']) != single_roll_no:
                continue
            
            # Calculate result (returns None if invalid)
            student_result = calculate_student_result(student_row, subjects)
            
            if student_result:
                valid_results.append(student_result)
            else:
                invalid_count += 1
                if not pd.isna(student_row.get('RollNo')):
                    errors.append(f"Invalid data for Roll No: {student_row.get('RollNo', 'Unknown')}")
            
            if progress_callback:
                progress_callback(idx + 1, len(students))
                
        except Exception as e:
            invalid_count += 1
    
    # Calculate positions for valid students
    if valid_results:
        valid_results = calculate_positions(valid_results)
        
        # Generate PDF
        pdf_filename = f"{class_name}_Results.pdf"
        pdf_path = os.path.join(output_subdir, pdf_filename)
        
        if generate_result_pdf(valid_results, class_name, pdf_path):
            return len(valid_results), invalid_count, errors
    
    return 0, invalid_count, errors


# ============================================================================
# GUI APPLICATION
# ============================================================================

class ResultSystemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("The Blessing School - Result System v3.0")
        self.root.geometry("800x700")
        
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.class_data_dir = os.path.join(self.base_dir, "CLASS_DATA")
        self.output_dir = os.path.join(self.base_dir, "OUTPUT")
        
        os.makedirs(self.class_data_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)
        
        self.current_subjects = None
        self.processing = False
        
        self.setup_ui()
    
    def setup_ui(self):
        """Create the GUI layout"""
        # Title
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill="x", padx=10, pady=10)
        
        title = ttk.Label(title_frame, 
                         text="🏫 The Blessing School - RESULT CARD GENERATOR v3.0 (Premium)", 
                         font=("Arial", 12, "bold"))
        title.pack()
        
        ttk.Separator(self.root).pack(fill="x", padx=10, pady=5)
        
        # Main content
        content = ttk.Frame(self.root)
        content.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Class selection
        ttk.Label(content, text="SELECT CLASS:").pack(anchor="w")
        self.class_var = tk.StringVar(value=CLASSES[0])
        class_combo = ttk.Combobox(content, textvariable=self.class_var, 
                                  values=CLASSES, state="readonly", width=30)
        class_combo.pack(anchor="w", pady=5)
        class_combo.bind("<<ComboboxSelected>>", self.on_class_selected)
        
        # Info
        ttk.Label(content, text="CLASS DETAILS:").pack(anchor="w", pady=(15, 5))
        self.info_text = tk.Text(content, height=6, width=70, state="disabled", 
                                font=("Courier", 9), bg="#f0f0f0")
        self.info_text.pack(pady=5)
        
        # Options
        ttk.Label(content, text="GENERATION OPTIONS:").pack(anchor="w", pady=(15, 5))
        
        self.option_var = tk.StringVar(value="all")
        ttk.Radiobutton(content, text="ALL STUDENTS", 
                       variable=self.option_var, value="all").pack(anchor="w")
        
        roll_frame = ttk.Frame(content)
        roll_frame.pack(anchor="w", pady=5)
        ttk.Radiobutton(roll_frame, text="SINGLE - Roll No:", 
                       variable=self.option_var, value="single").pack(side="left")
        self.roll_var = tk.StringVar()
        ttk.Entry(roll_frame, textvariable=self.roll_var, width=15).pack(side="left", padx=10)
        
        # Buttons
        btn_frame = ttk.Frame(content)
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text="📥 GENERATE PDFs", 
                  command=self.generate_pdfs).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="📂 OPEN OUTPUT", 
                  command=self.open_output).pack(side="left", padx=5)
        
        # Progress
        ttk.Label(content, text="PROGRESS:").pack(anchor="w", pady=(15, 5))
        self.progress = ttk.Progressbar(content, length=300, mode="determinate")
        self.progress.pack(anchor="w", pady=5)
        
        self.status_label = ttk.Label(content, text="Ready...", foreground="black")
        self.status_label.pack(anchor="w")
        
        # Features
        features = "✨ Features: 2 Cards/Page | Position Ranking | Auto-validation | Premium Design"
        ttk.Label(content, text=features, foreground="green", font=("Arial", 8)).pack(pady=10)
        
        self.update_class_info()
    
    def update_class_info(self):
        """Update class info display"""
        class_name = self.class_var.get()
        excel_file = os.path.join(self.class_data_dir, f"{class_name}.xlsx")
        
        self.info_text.config(state="normal")
        self.info_text.delete("1.0", "end")
        
        if os.path.exists(excel_file):
            result = load_excel_file(excel_file)
            if result:
                df, subjects = result
                self.current_subjects = subjects
                
                info = f"✅ File: {class_name}.xlsx\n"
                info += f"📊 Students: {len(df)}\n"
                info += f"📚 Subjects: {', '.join([s[0] for s in subjects])}\n"
                info += f"⚠️ Invalid entries will be skipped automatically"
                self.info_text.insert("end", info)
            else:
                self.info_text.insert("end", "❌ Invalid file format")
                self.current_subjects = None
        else:
            self.info_text.insert("end", f"❌ File not found: {class_name}.xlsx")
            self.current_subjects = None
        
        self.info_text.config(state="disabled")
    
    def on_class_selected(self, event=None):
        self.update_class_info()
    
    def generate_pdfs(self):
        """Generate PDFs"""
        if self.processing:
            messagebox.showwarning("Processing", "Already processing!")
            return
        
        class_name = self.class_var.get()
        
        if not self.current_subjects:
            messagebox.showerror("Error", "Invalid class file")
            return
        
        single_roll = None
        if self.option_var.get() == "single":
            try:
                single_roll = int(self.roll_var.get())
            except:
                messagebox.showerror("Error", "Invalid Roll Number")
                return
        
        self.processing = True
        thread = threading.Thread(target=self.generate_pdfs_thread, 
                                 args=(class_name, single_roll))
        thread.daemon = True
        thread.start()
    
    def generate_pdfs_thread(self, class_name, single_roll):
        """Background thread"""
        try:
            self.status_label.config(text="Processing...")
            
            def progress_callback(current, total):
                pct = int((current / total) * 100) if total > 0 else 0
                self.progress['value'] = pct
                self.status_label.config(text=f"Processing: {current}/{total}")
                self.root.update()
            
            generated, invalid, errors = process_class_to_pdf(
                class_name, self.class_data_dir, self.output_dir,
                single_roll, progress_callback
            )
            
            self.progress['value'] = 100
            
            output_file = os.path.join(self.output_dir, f"{class_name}_results", f"{class_name}_Results.pdf")
            
            if generated > 0:
                msg = f"✅ Generated: {generated} cards\n"
                if invalid > 0:
                    msg += f"⚠️ Skipped: {invalid} invalid entries\n"
                msg += f"\n📁 {output_file}"
                
                self.status_label.config(text=f"Success: {generated} cards")
                messagebox.showinfo("Success", msg)
            else:
                self.status_label.config(text="No valid students")
                messagebox.showerror("Error", "No valid students found")
        
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.processing = False
    
    def open_output(self):
        """Open output folder"""
        try:
            if sys.platform == "win32":
                os.startfile(self.output_dir)
        except:
            messagebox.showerror("Error", "Cannot open folder")


# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    root = tk.Tk()
    app = ResultSystemGUI(root)
    root.mainloop()