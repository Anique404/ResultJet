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
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
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

CLASSES = [
    "Playgroup (RED)", "Playgroup (BLUE)",
    "KG-1 (RED)", "KG-1 (BLUE)",
    "KG-2 (RED)", "KG-2 (BLUE)",
    "ONE", "TWO", "THREE", "FOUR", "FIVE",
    "SIXTH", "SEVEN", "EIGHT", "NINE", "TEN"
]

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
    fixed_columns = set(['RollNo', 'StudentName'])
    
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
# PDF GENERATION - MULTIPLE STUDENTS PER PAGE (A4 LANDSCAPE)
# ============================================================================

def create_result_card(student_result: Dict, class_name: str) -> Table:
    """
    A4 Landscape narrow form ke liye optimized card
    """
    card_elements = []
    
    # Logo and School Name in one row
    header_data = []
    
    # Check for logo
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=0.8*inch, height=0.8*inch)
        header_data.append([logo, Paragraph(SCHOOL_NAME, 
                           ParagraphStyle('School', fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER))])
        header_table = Table(header_data, colWidths=[0.8*inch, 6.7*inch])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (0, 0), 'RIGHT'),
            ('ALIGN', (1, 0), (1, 0), 'LEFT'),
        ]))
    else:
        header_data.append(['', Paragraph(SCHOOL_NAME, 
                           ParagraphStyle('School', fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER))])
        header_table = Table(header_data, colWidths=[0.8*inch, 6.7*inch])
    
    card_elements.append(header_table)
    card_elements.append(Spacer(1, 0.05*inch))
    
    # Student Details - Ek line mein
    details_data = [
        [f"Roll No: {student_result['roll_no']}   |   Class: {class_name}   |   {student_result['name']}"]
    ]
    
    details_table = Table(details_data, colWidths=[7.5*inch])
    details_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), 'Helvetica', 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f0f0f0')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    card_elements.append(details_table)
    card_elements.append(Spacer(1, 0.05*inch))
    
    # Subjects Table
    subject_data = [['Subject', 'Marks', 'Max', '%', 'Grade', 'Status']]
    for sub in student_result['subjects']:
        subject_data.append([
            sub['name'][:12],
            f"{sub['obtained']:.0f}",
            f"{sub['max']:.0f}",
            f"{sub['percentage']:.0f}%",
            sub['grade'],
            sub['status']
        ])
    
    subject_table = Table(subject_data, 
                         colWidths=[2.2*inch, 0.8*inch, 0.8*inch, 1.0*inch, 0.8*inch, 1.0*inch])
    subject_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), 'Helvetica', 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a1a4d')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f9f9f9')]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    card_elements.append(subject_table)
    card_elements.append(Spacer(1, 0.05*inch))
    
    # Overall Result
    status_color = colors.HexColor('#00aa00') if student_result['overall_status'] == 'PASS' else colors.HexColor('#cc0000')
    
    result_data = [[
        f"Total: {student_result['total_obtained']:.0f}/{student_result['total_max']:.0f}",
        f"Percentage: {student_result['overall_percentage']:.1f}%",
        f"Grade: {student_result['overall_grade']}",
        f"Result: {student_result['overall_status']}"
    ]]
    
    result_table = Table(result_data, colWidths=[1.8*inch, 1.8*inch, 1.8*inch, 1.8*inch])
    result_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), 'Helvetica', 8),
        ('FONT', (0, 0), (-1, -1), 'Helvetica-Bold', 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('TEXTCOLOR', (3, 0), (3, 0), status_color),
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#e6e6ff')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    card_elements.append(result_table)
    
    # Signature Line
    signature_data = [[f"Date: {datetime.now().strftime('%d-%m-%Y')}", "Principal Signature: _______________"]]
    signature_table = Table(signature_data, colWidths=[2.5*inch, 5.0*inch])
    signature_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), 'Helvetica', 7),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))
    card_elements.append(signature_table)
    
    # Card border
    card_table = Table([[e] for e in card_elements])
    card_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#1a1a4d')),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    
    return card_table


def generate_multi_result_pdf(students_results: List[Dict], class_name: str, output_path: str) -> bool:
    """
    A4 Landscape - 2 cards per page (vertical)
    """
    try:
        # A4 Landscape size
        doc = SimpleDocTemplate(output_path, 
                               pagesize=landscape(A4),
                               rightMargin=0.3*inch, 
                               leftMargin=0.3*inch,
                               topMargin=0.3*inch, 
                               bottomMargin=0.3*inch)
        
        elements = []
        
        for i in range(0, len(students_results), 2):
            student1 = students_results[i]
            student2 = students_results[i+1] if i+1 < len(students_results) else None
            
            # Card 1
            card1 = create_result_card(student1, class_name)
            elements.append(card1)
            elements.append(Spacer(1, 0.2*inch))
            
            # Card 2 (if exists)
            if student2:
                card2 = create_result_card(student2, class_name)
                elements.append(card2)
            
            # New page if more students left
            if i + 2 < len(students_results):
                elements.append(PageBreak())
        
        doc.build(elements)
        return True
        
    except Exception as e:
        print(f"Error in PDF generation: {e}")
        return False


# ============================================================================
# EXCEL FILE PROCESSING
# ============================================================================

def load_excel_file(file_path: str) -> Optional[Tuple[pd.DataFrame, List[Tuple[str, str, Optional[str]]]]]:
    """Load Excel file and detect subjects with max marks."""
    try:
        df = pd.read_excel(file_path)
        
        # Validate required columns
        required_cols = ['RollNo', 'StudentName']
        for col in required_cols:
            if col not in df.columns:
                print(f"Missing required column: {col}")
                return None
        
        # Detect subjects
        subjects = detect_subjects_and_max_marks(df)
        
        if not subjects:
            print("No subjects detected")
            return None
        
        return df, subjects
        
    except Exception as e:
        print(f"Error loading Excel file: {e}")
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
        error_messages.append(f"Failed to load Excel file: {class_name}.xlsx")
        return 0, 0, error_messages
    
    df, subjects = result
    
    # Create output subdirectory
    output_subdir = os.path.join(output_dir, f"{class_name}_results")
    os.makedirs(output_subdir, exist_ok=True)
    
    # Process students
    students = df.to_dict('records')
    all_results = []
    
    for idx, student_row in enumerate(students):
        try:
            # Filter by roll number if specified
            if single_roll_no and int(student_row['RollNo']) != single_roll_no:
                continue
            
            # Calculate result
            student_result = calculate_student_result(student_row, subjects)
            all_results.append(student_result)
            
            # Update progress
            if progress_callback:
                progress_callback(idx + 1, len(students))
                
        except Exception as e:
            total_failed += 1
            error_messages.append(f"Error processing student {student_row.get('RollNo', 'Unknown')}: {e}")
    
    # Generate single PDF with all results
    if all_results:
        pdf_filename = f"{class_name}_All_Results_Landscape.pdf"
        pdf_path = os.path.join(output_subdir, pdf_filename)
        
        if generate_multi_result_pdf(all_results, class_name, pdf_path):
            total_generated = len(all_results)
        else:
            total_failed = len(all_results)
            error_messages.append("Failed to generate PDF")
    
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
        self.class_var = tk.StringVar(value=CLASSES[0])
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
        
        # Note about output format
        note_frame = ttk.Frame(content_frame)
        note_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(note_frame, text="📄 Output: A4 Landscape | 2 Cards per page | Single PDF file", 
                 foreground="blue", font=("Arial", 8)).pack()
        
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
            pdf_file = os.path.join(output_subdir, f"{class_name}_All_Results_Landscape.pdf")
            
            if generated > 0:
                self.status_label.config(text=f"✅ SUCCESS: {generated} result cards generated")
                self.output_label.config(text=f"📁 OUTPUT: {pdf_file}")
                messagebox.showinfo("Success", 
                    f"✅ Successfully generated {generated} result cards in a single PDF!\n\n"
                    + (f"⚠️ {failed} cards failed\n" if failed > 0 else "") +
                    f"\n📁 Location: {pdf_file}")
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