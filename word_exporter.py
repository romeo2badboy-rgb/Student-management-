"""
Word Exporter Module - Advanced Version
Creates professional exam seating reports with visual seating maps.
"""

from typing import List, Optional, Dict
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from seating_algorithm import SeatingResult, Desk, Column


class WordExporter:
    """
    Advanced Word Document Exporter for Exam Seating.

    Creates professional reports with:
    - University header and department info
    - Formatted seating table
    - Visual seating map representation
    - Statistics summary
    - Signature section
    """

    # Color scheme for stages
    STAGE_COLORS = {
        "1st Stage": "DBEAFE",  # Light blue
        "2nd Stage": "DCFCE7",  # Light green
        "3rd Stage": "FEF3C7",  # Light yellow
    }

    DEFAULT_COLOR = "F3F4F6"  # Light gray

    def __init__(self):
        """Initialize the exporter."""
        self._document: Optional[Document] = None

    def export(
        self,
        result: SeatingResult,
        university_name: str,
        department_name: str,
        output_path: str,
        exam_title: str = "Exam Seating Arrangement",
        exam_date: Optional[str] = None,
        include_map: bool = True
    ) -> tuple[bool, str]:
        """
        Export seating arrangement to a Word document.

        Args:
            result: SeatingResult from the algorithm
            university_name: Name of the university
            department_name: Name of the department
            output_path: Path for the output file
            exam_title: Title for the document
            exam_date: Date of the exam (optional)
            include_map: Whether to include visual seating map

        Returns:
            Tuple of (success, message)
        """
        if not result or not result.desks:
            return False, "No seating data to export."

        if not output_path.lower().endswith('.docx'):
            output_path += '.docx'

        try:
            self._document = Document()
            self._setup_document()

            # Add document sections
            self._add_header(university_name, department_name, exam_title, exam_date)
            self._add_statistics(result)
            self._add_seating_table(result)

            if include_map:
                self._add_seating_map(result)

            self._add_legend()
            self._add_footer()

            self._document.save(output_path)
            return True, f"Successfully exported to {output_path}"

        except PermissionError:
            return False, f"Permission denied. Please close the file: {output_path}"
        except Exception as e:
            return False, f"Export error: {str(e)}"

    def _setup_document(self) -> None:
        """Configure document settings."""
        for section in self._document.sections:
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)
            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)

    def _add_header(
        self,
        university_name: str,
        department_name: str,
        exam_title: str,
        exam_date: Optional[str]
    ) -> None:
        """Add document header with university and exam info."""
        # University name
        uni_para = self._document.add_paragraph()
        uni_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        uni_run = uni_para.add_run(university_name)
        uni_run.bold = True
        uni_run.font.size = Pt(18)
        uni_run.font.color.rgb = RGBColor(30, 64, 175)  # Blue

        # Department
        dept_para = self._document.add_paragraph()
        dept_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dept_run = dept_para.add_run(f"Department of {department_name}")
        dept_run.bold = True
        dept_run.font.size = Pt(14)

        # Divider line
        divider = self._document.add_paragraph()
        divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
        divider.add_run("─" * 60)

        # Exam title
        title_para = self._document.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run(exam_title)
        title_run.bold = True
        title_run.font.size = Pt(16)

        # Date
        date_para = self._document.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_text = exam_date if exam_date else datetime.now().strftime("%B %d, %Y")
        date_run = date_para.add_run(f"Date: {date_text}")
        date_run.font.size = Pt(12)

        self._document.add_paragraph()

    def _add_statistics(self, result: SeatingResult) -> None:
        """Add statistics summary section."""
        stats_title = self._document.add_paragraph()
        stats_run = stats_title.add_run("Summary Statistics")
        stats_run.bold = True
        stats_run.font.size = Pt(12)

        # Create stats table
        stats = result.stats
        table = self._document.add_table(rows=2, cols=4)
        table.style = 'Table Grid'

        # Headers
        headers = ["Total Desks", "Total Students", "Cross-Stage Pairs", "Same-Stage Pairs"]
        values = [
            str(stats.get("total_desks", 0)),
            str(stats.get("total_students", 0)),
            str(stats.get("cross_stage_pairs", 0)),
            str(stats.get("same_stage_pairs", 0))
        ]

        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = header
            self._style_cell(cell, bold=True, bg_color="E5E7EB", align_center=True)

        for i, value in enumerate(values):
            cell = table.rows[1].cells[i]
            cell.text = value
            self._style_cell(cell, align_center=True)

        # Stage breakdown
        if "by_stage" in stats:
            self._document.add_paragraph()
            stage_para = self._document.add_paragraph()
            stage_para.add_run("Students by Stage: ").bold = True
            stage_info = [f"{stage}: {count}" for stage, count in stats["by_stage"].items()]
            stage_para.add_run(" | ".join(stage_info))

        # Warnings
        if result.warnings:
            self._document.add_paragraph()
            warn_para = self._document.add_paragraph()
            warn_run = warn_para.add_run("⚠ Warnings: ")
            warn_run.bold = True
            warn_run.font.color.rgb = RGBColor(180, 83, 9)  # Orange
            warn_para.add_run("; ".join(result.warnings))

        self._document.add_paragraph()

    def _add_seating_table(self, result: SeatingResult) -> None:
        """Add the main seating arrangement table."""
        # Section title
        title_para = self._document.add_paragraph()
        title_run = title_para.add_run("Seating Arrangement")
        title_run.bold = True
        title_run.font.size = Pt(14)

        # Create table
        table = self._document.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Set column widths
        widths = [Cm(1.2), Cm(1.5), Cm(1), Cm(5), Cm(5), Cm(2.5)]
        for i, width in enumerate(widths):
            for cell in table.columns[i].cells:
                cell.width = width

        # Header row
        headers = ["Desk", "Column", "Row", "Student A", "Student B", "Status"]
        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = header
            self._style_cell(cell, bold=True, bg_color="1E40AF", text_color="FFFFFF", align_center=True)

        # Data rows
        for desk in result.desks:
            row = table.add_row()

            # Desk number
            row.cells[0].text = str(desk.number)
            self._style_cell(row.cells[0], align_center=True, bold=True)

            # Column
            row.cells[1].text = desk.column.value
            self._style_cell(row.cells[1], align_center=True)

            # Row
            row.cells[2].text = str(desk.row)
            self._style_cell(row.cells[2], align_center=True)

            # Student A
            if desk.student_a:
                row.cells[3].text = f"{desk.student_a.name}\n({desk.student_a.stage})"
                color = self.STAGE_COLORS.get(desk.student_a.stage, self.DEFAULT_COLOR)
                self._style_cell(row.cells[3], bg_color=color)
            else:
                row.cells[3].text = "-"
                self._style_cell(row.cells[3], align_center=True)

            # Student B
            if desk.student_b:
                row.cells[4].text = f"{desk.student_b.name}\n({desk.student_b.stage})"
                color = self.STAGE_COLORS.get(desk.student_b.stage, self.DEFAULT_COLOR)
                self._style_cell(row.cells[4], bg_color=color)
            else:
                row.cells[4].text = "-"
                self._style_cell(row.cells[4], align_center=True)

            # Status
            if desk.is_cross_stage:
                row.cells[5].text = "✓ Cross-Stage"
                self._style_cell(row.cells[5], bg_color="D1FAE5", align_center=True)
            elif desk.is_full:
                row.cells[5].text = "⚠ Same Stage"
                self._style_cell(row.cells[5], bg_color="FEF3C7", align_center=True)
            else:
                row.cells[5].text = "Single"
                self._style_cell(row.cells[5], bg_color="E5E7EB", align_center=True)

        self._document.add_paragraph()

    def _add_seating_map(self, result: SeatingResult) -> None:
        """Add visual seating map representation."""
        # Page break before map
        self._document.add_page_break()

        # Title
        title_para = self._document.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run("VISUAL SEATING MAP")
        title_run.bold = True
        title_run.font.size = Pt(16)

        # Orientation note
        orient_para = self._document.add_paragraph()
        orient_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        orient_para.add_run("[ FRONT OF EXAMINATION HALL ]").font.size = Pt(10)

        self._document.add_paragraph()

        # Column headers
        col_header = self._document.add_paragraph()
        col_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        col_header.add_run("        RIGHT COLUMN          |         MIDDLE COLUMN         |          LEFT COLUMN        ")

        self._document.add_paragraph()

        # Group desks by row
        desks_by_row: Dict[int, Dict[str, Desk]] = {}
        for desk in result.desks:
            if desk.row not in desks_by_row:
                desks_by_row[desk.row] = {}
            desks_by_row[desk.row][desk.column.value] = desk

        # Create visual map table
        if desks_by_row:
            max_row = max(desks_by_row.keys())
            table = self._document.add_table(rows=max_row, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            # Set column widths
            for col in table.columns:
                for cell in col.cells:
                    cell.width = Cm(5.5)

            for row_num in range(1, max_row + 1):
                row_data = desks_by_row.get(row_num, {})
                table_row = table.rows[row_num - 1]

                for col_idx, col_name in enumerate(["Right", "Middle", "Left"]):
                    cell = table_row.cells[col_idx]
                    desk = row_data.get(col_name)

                    if desk:
                        self._format_map_cell(cell, desk)
                    else:
                        cell.text = ""
                        self._style_cell(cell, bg_color="F9FAFB")

        self._document.add_paragraph()

        # Back of hall note
        back_para = self._document.add_paragraph()
        back_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        back_para.add_run("[ BACK OF EXAMINATION HALL ]").font.size = Pt(10)

    def _format_map_cell(self, cell, desk: Desk) -> None:
        """Format a cell in the visual map."""
        cell.text = ""

        # Add desk number
        p1 = cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        desk_run = p1.add_run(f"Desk {desk.number}")
        desk_run.bold = True
        desk_run.font.size = Pt(9)

        # Add students
        if desk.student_a:
            p2 = cell.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            name_a = desk.student_a.name[:15] + "..." if len(desk.student_a.name) > 15 else desk.student_a.name
            run_a = p2.add_run(f"A: {name_a}")
            run_a.font.size = Pt(8)

        if desk.student_b:
            p3 = cell.add_paragraph()
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            name_b = desk.student_b.name[:15] + "..." if len(desk.student_b.name) > 15 else desk.student_b.name
            run_b = p3.add_run(f"B: {name_b}")
            run_b.font.size = Pt(8)

        # Color based on status
        if desk.is_cross_stage:
            self._set_cell_shading(cell, "D1FAE5")  # Green - good
        elif desk.is_full:
            self._set_cell_shading(cell, "FEF3C7")  # Yellow - warning
        else:
            self._set_cell_shading(cell, "E5E7EB")  # Gray - single

        # Set vertical alignment
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _add_legend(self) -> None:
        """Add color legend."""
        self._document.add_paragraph()

        legend_title = self._document.add_paragraph()
        legend_title.add_run("Legend:").bold = True

        # Stage colors
        for stage, color in self.STAGE_COLORS.items():
            p = self._document.add_paragraph()
            p.add_run(f"  ■ {stage}").font.size = Pt(10)

        # Status colors
        self._document.add_paragraph()
        status_para = self._document.add_paragraph()
        status_para.add_run("Status: ").bold = True
        status_para.add_run("Green = Cross-Stage (Good) | Yellow = Same Stage (Warning) | Gray = Single Student")

    def _add_footer(self) -> None:
        """Add footer with signatures and timestamp."""
        self._document.add_paragraph()
        self._document.add_paragraph()

        # Signature table
        sig_table = self._document.add_table(rows=3, cols=3)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        sig_table.cell(0, 0).text = "Prepared by:"
        sig_table.cell(0, 1).text = "Verified by:"
        sig_table.cell(0, 2).text = "Approved by:"

        sig_table.cell(1, 0).text = ""
        sig_table.cell(1, 1).text = ""
        sig_table.cell(1, 2).text = ""

        sig_table.cell(2, 0).text = "_" * 20
        sig_table.cell(2, 1).text = "_" * 20
        sig_table.cell(2, 2).text = "_" * 20

        for row in sig_table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Timestamp
        self._document.add_paragraph()
        time_para = self._document.add_paragraph()
        time_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        time_run = time_para.add_run(
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        time_run.font.size = Pt(8)
        time_run.italic = True

    def _style_cell(
        self,
        cell,
        bold: bool = False,
        bg_color: Optional[str] = None,
        text_color: Optional[str] = None,
        align_center: bool = False
    ) -> None:
        """Apply styling to a table cell."""
        for paragraph in cell.paragraphs:
            if align_center:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            for run in paragraph.runs:
                if bold:
                    run.bold = True
                if text_color:
                    run.font.color.rgb = RGBColor.from_string(text_color)

        if bg_color:
            self._set_cell_shading(cell, bg_color)

    def _set_cell_shading(self, cell, color: str) -> None:
        """Set background color for a table cell."""
        shading = parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="{color}" w:val="clear"/>'
        )
        cell._tc.get_or_add_tcPr().append(shading)
