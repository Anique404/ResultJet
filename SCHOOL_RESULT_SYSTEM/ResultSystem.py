"""
The Blessing School - RESULT CARD GENERATION SYSTEM v2.0
=========================================================
Desktop Application for generating PDF result cards from Excel files.
Supports KG and classes 1st to 10th.

Author: School Administration System
Date: 2026
"""

import sys
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import threading
from typing import List, Dict, Tuple, Optional
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import math
import subprocess
import webbrowser

# ============================================================================
# CONFIGURATION
# ============================================================================

PASS_THRESHOLD = 33  # Minimum percentage required to pass a subject
SCHOOL_NAME = "The Blessing School"
SCHOOL_CITY = "Karachi, Pakistan"

CLASSES = ["KG", "1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th"]

# ============================================================================
# CORE FUNCTIONS - SUBJECT & PERCENTAGE CALCULATION
# ============================================================================

def detect_subjects_and_max_marks(df: pd.DataFrame) -> List[Tuple[str, str, Optional[str]]]:
    """
    CRITICAL FUNCTION: Detects subjects and their max marks from Excel columns.
    
    Rules:
    - Columns ending with '_Max' are MAX MARKS columns
    - Subject name = column_name.replace('_Max', '')
    - If no _Max column found for a subject → DEFAULT MAX = 100
    
    Args:
        df: DataFrame from Excel file
        
    Returns:
        List of tuples: [(subject_name, marks_column, max_marks_column), ...]
    """
    fixed_columns = set(['RollNo', 'StudentName', 'FatherName'])
    
    max_mark_columns = {}
    marks_columns = set()
    
    # Find all _Max columns
    for col in df.columns:
        if col.endswith('_Max'):
            subject = col.replace('_Max', '')
            max_mark_columns[subject] = col
        elif col not in fixed_columns:
            marks_columns.add(col)
    
    # Build subjects list
    subjects = []
    for col in df.columns:
        if col in fixed_columns:
            continue
        
        # Skip if this is a _Max column (already processed)
        if col.endswith('_Max'):
            continue
            
        # Check if there's a corresponding _Max column
        max_col = col + '_Max'
        if max_col in df.columns:
            subjects.append((col, col, max_col))
        else:
            # Default max = 100
            subjects.append((col, col, None))
    
    return subjects


def calculate_subject_percentage(obtained: float, max_marks: float) -> float:
    """Calculate percentage: (obtained / max) * 100, rounded to 2 decimals"""
    if max_marks == 0:
        return 0.0
    percentage = (obtained / max_marks) * 100
    return round(percentage, 2)


def is_subject_pass(percentage: float) -> bool:
    """Return True if percentage >= 33%, else False"""
    return percentage >= PASS_THRESHOLD


def get_grade(percentage: float) -> str:
    """
    Assign grade based on percentage.
    A+(90+), A(80-89), B(70-79), C(60-69), D(50-59), E(33-49), F(<33)
    """
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
    """
    CRITICAL: Determine PASS/FAIL based on subject percentages.
    FAIL if ANY subject is below 33%.
    PASS only if ALL subjects are 33% or above.
    """
    for percentage in subject_percentages:
        if percentage < PASS_THRESHOLD:
            return False  # FAIL
    return True  # PASS


# ============================================================================
# STUDENT RESULT CALCULATION
# ============================================================================

def calculate_student_result(student_row: pd.Series, subjects: List[Tuple[str, str, Optional[str]]]) -> Dict:
    """
    Calculate complete result for a student.
    """
    # allow student_row to be a dict or a pandas Series
    def _has(col):
        try:
            return col in student_row.index
        except Exception:
            return col in student_row

    def _get(col):
        try:
            return student_row[col]
        except Exception:
            return student_row.get(col)

    result = {
        'roll_no': int(_get('RollNo')),
        'name': _get('StudentName'),
        'father_name': _get('FatherName'),
        'subjects': [],
        'total_obtained': 0,
        'total_max': 0,
        'failed_subjects': []
    }
    
    subject_percentages = []
    
    for subject_name, marks_col, max_col in subjects:
        try:
            # Ensure marks column exists
            if not _has(marks_col):
                raise KeyError(f"Subject '{subject_name}' has marks column missing")

            obtained_raw = _get(marks_col)
            obtained = float(obtained_raw) if obtained_raw is not None and obtained_raw != '' else 0.0

            # Determine max marks
            if max_col and _has(max_col):
                max_marks_raw = _get(max_col)
                max_marks = float(max_marks_raw) if max_marks_raw is not None and max_marks_raw != '' else 100.0
            else:
                max_marks = 100.0  # DEFAULT MAX

            # Validate marks
            if obtained > max_marks:
                # Cap at max and record a warning (handled upstream if needed)
                obtained = max_marks

            # Calculate percentage
            percentage = calculate_subject_percentage(obtained, max_marks)
            subject_percentages.append(percentage)

            # Determine pass/fail for subject
            is_pass = is_subject_pass(percentage)
            status = "PASS" if is_pass else "FAIL"

            if not is_pass:
                result['failed_subjects'].append(f"{subject_name} ({percentage}%)")

            # Get grade
            grade = get_grade(percentage)

            # Add to subjects
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

        except Exception as e:
            # Append an error note to failed_subjects for visibility
            result['failed_subjects'].append(f"{subject_name} (ERROR: {e})")
            continue
    
    # Calculate overall results
    if result['total_max'] > 0:
        result['overall_percentage'] = calculate_subject_percentage(result['total_obtained'], result['total_max'])
        result['overall_grade'] = get_grade(result['overall_percentage'])
        result['overall_status'] = "PASS" if determine_overall_result(subject_percentages) else "FAIL"
    else:
        result['overall_percentage'] = 0
        result['overall_grade'] = "F"
        result['overall_status'] = "FAIL"
    
    return result


# ============================================================================
# PDF GENERATION
# ============================================================================

def generate_result_card_pdf(student_result: Dict, class_name: str, output_path: str) -> bool:
    """
    Generate a PDF result card for a student.
    """
    try:
        # Create PDF
        doc = SimpleDocTemplate(output_path, pagesize=A4,
                               rightMargin=0.5*inch, leftMargin=0.5*inch,
                               topMargin=0.5*inch, bottomMargin=0.5*inch)
        
        elements = []
        
        # Header
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=getSampleStyleSheet()['Heading1'],
            fontSize=14,
            textColor=colors.HexColor('#1a1a4d'),
            spaceAfter=2,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )
        
        elements.append(Paragraph(SCHOOL_NAME, header_style))
        
        subheader_style = ParagraphStyle(
            'Subheader',
            parent=getSampleStyleSheet()['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#333333'),
            spaceAfter=6,
            alignment=TA_CENTER,
            fontName='Helvetica'
        )
        
        elements.append(Paragraph("ANNUAL EXAMINATION RESULT", subheader_style))
        elements.append(Spacer(1, 0.15*inch))
        
        # Student Details
        details_data = [
            ['ROLL NO:', str(student_result['roll_no'])],
            ['STUDENT NAME:', student_result['name']],
            ['FATHER NAME:', student_result['father_name']],
            ['CLASS:', f"{class_name} {chr(65)}"],
        ]
        
        details_table = Table(details_data, colWidths=[2.2*inch, 4.3*inch])
        details_table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'Helvetica', 10),
            ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#1a1a4d')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        
        elements.append(details_table)
        elements.append(Spacer(1, 0.15*inch))
        
        # Subject-wise results table
        subject_data = [['SUBJECT', 'MARKS', 'MAX', '%', 'GRADE', 'STATUS']]
        
        for subject in student_result['subjects']:
            marks_str = f"{subject['obtained']:.0f}" if subject['obtained'] == int(subject['obtained']) else f"{subject['obtained']:.2f}"
            max_str = f"{subject['max']:.0f}" if subject['max'] == int(subject['max']) else f"{subject['max']:.2f}"
            
            subject_data.append([
                subject['name'][:20],
                marks_str,
                max_str,
                f"{subject['percentage']:.2f}%",
                subject['grade'],
                subject['status']
            ])
        
        # Add total row
        total_obtained_str = f"{student_result['total_obtained']:.0f}" if student_result['total_obtained'] == int(student_result['total_obtained']) else f"{student_result['total_obtained']:.2f}"
        total_max_str = f"{student_result['total_max']:.0f}" if student_result['total_max'] == int(student_result['total_max']) else f"{student_result['total_max']:.2f}"
        
        subject_data.append([
            'TOTAL',
            total_obtained_str,
            total_max_str,
            f"{student_result['overall_percentage']:.2f}%",
            student_result['overall_grade'],
            ''
        ])
        
        # Create subject table
        subject_table = Table(subject_data, colWidths=[2.0*inch, 1.0*inch, 0.8*inch, 1.0*inch, 0.8*inch, 0.9*inch])
        subject_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a1a4d')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 9),
            ('FONT', (0, 1), (-1, -2), 'Helvetica', 9),
            ('FONT', (0, -1), (-1, -1), 'Helvetica-Bold', 9),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#e6e6ff')),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f5f5f5')]),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        elements.append(subject_table)
        elements.append(Spacer(1, 0.2*inch))
        
        # Overall Result Section
        result_color = colors.HexColor('#00aa00') if student_result['overall_status'] == 'PASS' else colors.HexColor('#cc0000')
        
        overall_data = [
            ['TOTAL MARKS:', f"{student_result['total_obtained']:.0f} / {student_result['total_max']:.0f}"],
            ['OVERALL %:', f"{student_result['overall_percentage']:.2f}%"],
            ['GRADE:', student_result['overall_grade']],
            ['RESULT:', f"{student_result['overall_status']}"]
        ]
        
        # Add failed subjects info if applicable
        if student_result['failed_subjects']:
            failed_text = ", ".join(student_result['failed_subjects'])
            overall_data.append(['FAILED IN:', failed_text])
        
        overall_table = Table(overall_data, colWidths=[2.5*inch, 4.0*inch])
        overall_table.setStyle(TableStyle([
            ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
            ('FONT', (1, 0), (1, -2), 'Helvetica', 10),
            ('FONT', (1, -1), (1, -1), 'Helvetica-Bold', 12),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#1a1a4d')),
            ('TEXTCOLOR', (1, -1), (1, -1), result_color),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#ffffcc')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ]))
        
        elements.append(overall_table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Footer with signature space
        footer_data = [
            ['PRINCIPAL SIGNATURE: _______________', f"DATE: {datetime.now().strftime('%d-%m-%Y')}"]
        ]
        
        footer_table = Table(footer_data, colWidths=[3.5*inch, 2.5*inch])
        footer_table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'Helvetica', 9),
            ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
            ('TOPPADDING', (0, 0), (-1, -1), 20),
        ]))
        
        elements.append(footer_table)
        
        # Build PDF
        doc.build(elements)
        return True
        
    except Exception as e:
        return False


# ============================================================================
# EXCEL FILE PROCESSING
# ============================================================================

def load_excel_file(file_path: str) -> Optional[Tuple[pd.DataFrame, List[Tuple[str, str, Optional[str]]]]]:
    """Load Excel file and detect subjects with max marks."""
    try:
        df = pd.read_excel(file_path)
        
        # Validate required columns
        required_cols = ['RollNo', 'StudentName', 'FatherName']
        for col in required_cols:
            if col not in df.columns:
                return None
        
        # Detect subjects
        subjects = detect_subjects_and_max_marks(df)
        
        if not subjects:
            return None
        
        return df, subjects
        
    except Exception as e:
        return None


def process_class_to_pdf(class_name: str, class_data_dir: str, output_dir: str, 
                         single_roll_no: Optional[int] = None, 
                         progress_callback=None) -> Tuple[int, int, List[str]]:
    """Process all students in a class and generate PDFs."""
    total_generated = 0
    total_failed = 0
    error_messages = []
    
    # Find Excel file
    excel_file = os.path.join(class_data_dir, f"{class_name}.xlsx")
    if not os.path.exists(excel_file):
        error_messages.append(f"Excel file not found: {class_name}.xlsx")
        return 0, 0, error_messages
    
    # Load Excel file
    result = load_excel_file(excel_file)
    if result is None:
        return 0, 0, ["Failed to load Excel file"]
    
    df, subjects = result
    
    # Create output subdirectory
    output_subdir = os.path.join(output_dir, f"{class_name}_results")
    os.makedirs(output_subdir, exist_ok=True)
    
    # Process students
    students = df.to_dict('records')
    
    for idx, student_row in enumerate(students):
        try:
            # Filter by roll number if specified
            if single_roll_no and int(student_row['RollNo']) != single_roll_no:
                continue
            
            # Calculate result
            student_result = calculate_student_result(student_row, subjects)
            
            # Generate PDF
            pdf_filename = f"{student_result['roll_no']:03d}_{student_result['name'].replace(' ', '_')}_{class_name}_Result.pdf"
            pdf_path = os.path.join(output_subdir, pdf_filename)
            
            if generate_result_card_pdf(student_result, class_name, pdf_path):
                total_generated += 1
            else:
                total_failed += 1
            
            # Update progress
            if progress_callback:
                progress_callback(idx + 1, len(students))
                
        except Exception as e:
            total_failed += 1
    
    return total_generated, total_failed, error_messages


# ============================================================================
# GUI APPLICATION (Tkinter-based)
# ============================================================================

class ResultSystemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("School Result System - The Blessing School")
        self.root.geometry("800x700")
        
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.class_data_dir = os.path.join(self.base_dir, "CLASS_DATA")
        self.output_dir = os.path.join(self.base_dir, "OUTPUT")
        
        # Create directories if needed
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
        
        title = ttk.Label(title_frame, text="🏫 The Blessing School - RESULT CARD GENERATOR v2.0", 
                         font=("Arial", 12, "bold"))
        title.pack()
        
        # Separator
        separator = ttk.Separator(self.root, orient="horizontal")
        separator.pack(fill="x", padx=10, pady=5)
        
        # Main content frame
        content_frame = ttk.Frame(self.root)
        content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Class selection
        ttk.Label(content_frame, text="SELECT CLASS:").pack(anchor="w")
        self.class_var = tk.StringVar(value="10th")
        class_combo = ttk.Combobox(content_frame, textvariable=self.class_var, 
                                  values=CLASSES, state="readonly", width=30)
        class_combo.pack(anchor="w", pady=5)
        class_combo.bind("<<ComboboxSelected>>", self.on_class_selected)
        
        # Class info
        ttk.Label(content_frame, text="CLASS DETAILS:").pack(anchor="w", pady=(15, 5))
        self.info_text = tk.Text(content_frame, height=5, width=70, state="disabled", 
                                font=("Courier", 9), bg="#f0f0f0")
        self.info_text.pack(pady=5)
        
        # Options
        ttk.Label(content_frame, text="GENERATION OPTIONS:").pack(anchor="w", pady=(15, 5))
        
        self.option_var = tk.StringVar(value="all")
        ttk.Radiobutton(content_frame, text="ALL STUDENTS - Generate all result cards", 
                       variable=self.option_var, value="all").pack(anchor="w")
        
        roll_frame = ttk.Frame(content_frame)
        roll_frame.pack(anchor="w", pady=5)
        ttk.Radiobutton(roll_frame, text="SINGLE STUDENT - Roll Number:", 
                       variable=self.option_var, value="single").pack(side="left")
        self.roll_var = tk.StringVar()
        roll_entry = ttk.Entry(roll_frame, textvariable=self.roll_var, width=15)
        roll_entry.pack(side="left", padx=10)
        
        # Buttons
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(pady=15)
        
        ttk.Button(button_frame, text="📥 GENERATE PDFs", 
                  command=self.generate_pdfs).pack(side="left", padx=5)
        ttk.Button(button_frame, text="📂 OPEN OUTPUT FOLDER", 
                  command=self.open_output).pack(side="left", padx=5)
        
        # Progress
        ttk.Label(content_frame, text="PROGRESS:").pack(anchor="w", pady=(15, 5))
        self.progress = ttk.Progressbar(content_frame, length=300, mode="determinate")
        self.progress.pack(anchor="w", pady=5)
        
        self.status_label = ttk.Label(content_frame, text="Ready...", foreground="black")
        self.status_label.pack(anchor="w")
        
        self.output_label = ttk.Label(content_frame, text="", foreground="green", 
                                     font=("Courier", 8))
        self.output_label.pack(anchor="w", pady=5)
        
        # Warning
        warning_frame = ttk.Frame(content_frame)
        warning_frame.pack(fill="x", pady=(15, 0))
        ttk.Label(warning_frame, text="⚠️ Pass Criteria: 33% in EACH subject required for PASS", 
                 foreground="orange", font=("Arial", 9)).pack()
        
        # Initial class info
        self.update_class_info()
    
    def update_class_info(self):
        """Update class information display"""
        class_name = self.class_var.get()
        excel_file = os.path.join(self.class_data_dir, f"{class_name}.xlsx")
        
        self.info_text.config(state="normal")
        self.info_text.delete("1.0", "end")
        
        if os.path.exists(excel_file):
            result = load_excel_file(excel_file)
            if result:
                df, subjects = result
                self.current_subjects = subjects
                
                info = f"✅ File found: {class_name}.xlsx\n"
                info += f"   • Total Students: {len(df)}\n"
                info += f"   • Subjects Found: {len(subjects)}\n"
                info += f"   • Subjects: {', '.join([s[0] for s in subjects])}\n"
                info += f"   • Max marks detected from _Max columns\n"
                info += f"   • Pass marks: 33% in each subject"
                self.info_text.insert("end", info)
            else:
                self.info_text.insert("end", "⚠️ Error loading class file")
                self.current_subjects = None
        else:
            info = f"❌ File not found: CLASS_DATA/{class_name}.xlsx\n"
            info += f"   Please create this file with the required format"
            self.info_text.insert("end", info)
            self.current_subjects = None
        
        self.info_text.config(state="disabled")
    
    def on_class_selected(self, event=None):
        """Handle class selection change"""
        self.update_class_info()
    
    def generate_pdfs(self):
        """Generate PDFs"""
        if self.processing:
            messagebox.showwarning("Processing", "Already processing!")
            return
        
        class_name = self.class_var.get()
        
        if not self.current_subjects:
            messagebox.showerror("Error", 
                f"Excel file not found or invalid:\n{self.class_data_dir}/{class_name}.xlsx")
            return
        
        single_roll = None
        if self.option_var.get() == "single":
            try:
                single_roll = int(self.roll_var.get())
            except:
                messagebox.showerror("Error", "Please enter a valid Roll Number")
                return
        
        # Run generation in separate thread
        self.processing = True
        thread = threading.Thread(target=self.generate_pdfs_thread, 
                                 args=(class_name, single_roll))
        thread.daemon = True
        thread.start()
    
    def generate_pdfs_thread(self, class_name, single_roll):
        """Generate PDFs in background thread"""
        try:
            self.status_label.config(text="Starting generation...")
            self.root.update()
            
            def progress_callback(current, total):
                pct = int((current / total) * 100) if total > 0 else 0
                self.progress['value'] = pct
                self.status_label.config(text=f"Processing: {current}/{total} students...")
                self.root.update()
            
            generated, failed, errors = process_class_to_pdf(
                class_name,
                self.class_data_dir,
                self.output_dir,
                single_roll,
                progress_callback
            )
            
            self.progress['value'] = 100
            
            output_subdir = os.path.join(self.output_dir, f"{class_name}_results")
            
            if generated > 0:
                self.status_label.config(text=f"✅ SUCCESS: {generated} result cards generated")
                self.output_label.config(text=f"📁 OUTPUT: {output_subdir}")
                messagebox.showinfo("Success", 
                    f"✅ Successfully generated {generated} result cards!\n\n"
                    + (f"⚠️ {failed} cards failed\n" if failed > 0 else "") +
                    f"\n📁 Location: {output_subdir}")
            else:
                self.status_label.config(text="❌ No result cards generated")
                messagebox.showerror("Error", "Failed to generate result cards")
        
        except Exception as e:
            self.status_label.config(text=f"❌ ERROR: {str(e)}")
            messagebox.showerror("Error", f"Error: {str(e)}")
        
        finally:
            self.processing = False
    
    def open_output(self):
        """Open output folder"""
        try:
            if sys.platform == "win32":
                os.startfile(self.output_dir)
            else:
                subprocess.Popen(['xdg-open', self.output_dir])
        except:
            messagebox.showerror("Error", "Could not open folder")


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point"""
    root = tk.Tk()
    gui = ResultSystemGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
