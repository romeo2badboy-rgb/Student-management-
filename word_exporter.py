"""
Word Exporter Module
Handles exporting seating arrangements to formatted Word documents.
"""

from typing import List, Optional
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from seating_algorithm import DeskAssignment


class WordExporter:
    """
    Exports seating arrangements to professionally formatted Word documents.
    Creates official-looking exam seating forms.
    """

    def __init__(self):
        """Initialize the Word exporter."""
        self._document: Optional[Document] = None

    def export(
        self,
        assignments: List[DeskAssignment],
        department_name: str,
        output_path: str,
        exam_title: str = "Exam Seating Arrangement",
        exam_date: Optional[str] = None
    ) -> tuple[bool, str]:
        """
        Export seating assignments to a Word document.

        Args:
            assignments: List of desk assignments
            department_name: Name of the department
            output_path: Path for the output file
            exam_title: Title for the exam document
            exam_date: Date of the exam (optional)

        Returns:
            Tuple of (success: bool, message: str)
        """
        if not assignments:
            return False, "No seating assignments to export."

        if not output_path.lower().endswith('.docx'):
            output_path += '.docx'

        try:
            self._document = Document()
            self._setup_document()
            self._add_header(department_name, exam_title, exam_date)
            self._add_seating_table(assignments)
            self._add_footer(assignments)

            self._document.save(output_path)
            return True, f"Successfully exported to {output_path}"

        except PermissionError:
            return False, f"Permission denied. Please close the file if it's open: {output_path}"
        except Exception as e:
            return False, f"Error exporting document: {str(e)}"

    def _setup_document(self) -> None:
        """Configure document settings and margins."""
        sections = self._document.sections
        for section in sections:
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)
            section.left_margin = Cm(2)
            section.right_margin = Cm(2)
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)

    def _add_header(
        self,
        department_name: str,
        exam_title: str,
        exam_date: Optional[str]
    ) -> None:
        """Add the document header with title and department info."""
        # Main title
        title = self._document.add_heading(exam_title, level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Department name
        dept_para = self._document.add_paragraph()
        dept_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dept_run = dept_para.add_run(f"Department: {department_name}")
        dept_run.bold = True
        dept_run.font.size = Pt(14)

        # Date
        date_para = self._document.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_text = exam_date if exam_date else datetime.now().strftime("%B %d, %Y")
        date_run = date_para.add_run(f"Date: {date_text}")
        date_run.font.size = Pt(12)

        # Add some spacing
        self._document.add_paragraph()

    def _add_seating_table(self, assignments: List[DeskAssignment]) -> None:
        """Add the main seating arrangement table."""
        # Table header
        table_title = self._document.add_paragraph()
        table_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = table_title.add_run("Seating Arrangement")
        title_run.bold = True
        title_run.font.size = Pt(12)

        # Create table with 6 columns
        table = self._document.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Set column widths
        widths = [Cm(1.5), Cm(2), Cm(4.5), Cm(2.5), Cm(4.5), Cm(2.5)]
        for i, width in enumerate(widths):
            table.columns[i].width = width

        # Add header row
        header_cells = table.rows[0].cells
        headers = ["Desk #", "Column", "Student A", "Stage A", "Student B", "Stage B"]

        for i, header in enumerate(headers):
            header_cells[i].text = header
            # Make header bold
            for paragraph in header_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Add shading to header
            self._set_cell_shading(header_cells[i], "D9E2F3")

        # Add data rows
        for assignment in assignments:
            row_cells = table.add_row().cells

            row_cells[0].text = str(assignment.desk_number)
            row_cells[1].text = assignment.column
            row_cells[2].text = assignment.student_a
            row_cells[3].text = assignment.student_a_stage
            row_cells[4].text = assignment.student_b if assignment.student_b else "-"
            row_cells[5].text = assignment.student_b_stage if assignment.student_b_stage else "-"

            # Center align desk number and column
            for j in [0, 1, 3, 5]:
                for paragraph in row_cells[j].paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Highlight if same stage (violation of rule)
            if assignment.student_b and assignment.student_a_stage == assignment.student_b_stage:
                for cell in row_cells:
                    self._set_cell_shading(cell, "FFE699")  # Yellow warning

    def _set_cell_shading(self, cell, color: str) -> None:
        """Set background color for a table cell."""
        shading = parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="{color}" w:val="clear"/>'
        )
        cell._tc.get_or_add_tcPr().append(shading)

    def _add_footer(self, assignments: List[DeskAssignment]) -> None:
        """Add footer with statistics and signature lines."""
        self._document.add_paragraph()

        # Statistics
        total_students = sum(
            1 + (1 if a.student_b else 0) for a in assignments
        )
        cross_stage = sum(
            1 for a in assignments
            if a.student_b and a.student_a_stage != a.student_b_stage
        )

        stats_para = self._document.add_paragraph()
        stats_para.add_run("Summary: ").bold = True
        stats_para.add_run(
            f"Total Desks: {len(assignments)} | "
            f"Total Students: {total_students} | "
            f"Cross-Stage Pairs: {cross_stage}"
        )

        # Add spacing
        self._document.add_paragraph()
        self._document.add_paragraph()

        # Signature lines
        sig_table = self._document.add_table(rows=2, cols=2)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        sig_table.cell(0, 0).text = "Prepared by:"
        sig_table.cell(0, 1).text = "Approved by:"
        sig_table.cell(1, 0).text = "_" * 30
        sig_table.cell(1, 1).text = "_" * 30

        for row in sig_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Generation timestamp
        self._document.add_paragraph()
        timestamp_para = self._document.add_paragraph()
        timestamp_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        timestamp_run = timestamp_para.add_run(
            f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        timestamp_run.font.size = Pt(8)
        timestamp_run.italic = True

    def export_by_column(
        self,
        assignments: List[DeskAssignment],
        department_name: str,
        output_path: str,
        exam_title: str = "Exam Seating Arrangement",
        exam_date: Optional[str] = None
    ) -> tuple[bool, str]:
        """
        Export seating arrangements organized by column (Left, Middle, Right).

        Args:
            assignments: List of desk assignments
            department_name: Name of the department
            output_path: Path for the output file
            exam_title: Title for the exam document
            exam_date: Date of the exam (optional)

        Returns:
            Tuple of (success: bool, message: str)
        """
        if not assignments:
            return False, "No seating assignments to export."

        if not output_path.lower().endswith('.docx'):
            output_path += '.docx'

        try:
            self._document = Document()
            self._setup_document()
            self._add_header(department_name, exam_title, exam_date)

            # Organize by column
            by_column = {"Left": [], "Middle": [], "Right": []}
            for assignment in assignments:
                by_column[assignment.column].append(assignment)

            # Add a section for each column
            for column in ["Left", "Middle", "Right"]:
                if by_column[column]:
                    self._add_column_section(column, by_column[column])

            self._add_footer(assignments)
            self._document.save(output_path)
            return True, f"Successfully exported to {output_path}"

        except PermissionError:
            return False, f"Permission denied. Please close the file if it's open: {output_path}"
        except Exception as e:
            return False, f"Error exporting document: {str(e)}"

    def _add_column_section(self, column_name: str, assignments: List[DeskAssignment]) -> None:
        """Add a section for a specific column."""
        # Column header
        col_header = self._document.add_paragraph()
        col_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
        header_run = col_header.add_run(f"📍 {column_name} Column")
        header_run.bold = True
        header_run.font.size = Pt(11)

        # Create table
        table = self._document.add_table(rows=1, cols=4)
        table.style = 'Table Grid'

        # Header row
        headers = ["Desk #", "Student A", "Student B", "Stages"]
        header_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            header_cells[i].text = header
            for paragraph in header_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_cell_shading(header_cells[i], "D9E2F3")

        # Data rows
        for assignment in assignments:
            row_cells = table.add_row().cells
            row_cells[0].text = str(assignment.desk_number)
            row_cells[1].text = assignment.student_a
            row_cells[2].text = assignment.student_b if assignment.student_b else "-"
            stages = f"{assignment.student_a_stage}"
            if assignment.student_b_stage:
                stages += f" / {assignment.student_b_stage}"
            row_cells[3].text = stages

            # Center align
            for paragraph in row_cells[0].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        self._document.add_paragraph()
