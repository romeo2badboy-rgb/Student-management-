"""
Word Exporter Module - Professional Version
Creates exam seating reports with desk-cell table layout.
"""

from typing import List, Optional, Dict
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml

from seating_algorithm import SeatingResult, Desk, Column


class WordExporter:
    """
    Professional Word Document Exporter for Exam Seating.

    Creates reports with:
    - University header and department info
    - Desk-cell table where each cell = one desk
    - Student Name (Bold), Stage (Smaller), Seat Number
    - Statistics summary and signature section
    """

    # Color scheme for stages
    STAGE_COLORS = {
        "1st Stage": "DBEAFE",  # Light blue
        "2nd Stage": "DCFCE7",  # Light green
        "3rd Stage": "FEF3C7",  # Light yellow
    }

    DEFAULT_COLOR = "F3F4F6"  # Light gray
    EMPTY_COLOR = "E5E7EB"    # Gray for empty seats

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
            self._add_detailed_table(result)

            if include_map:
                self._add_visual_seating_map(result)

            self._add_legend(result)
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
            section.page_width = Inches(11)  # Landscape for better desk view
            section.page_height = Inches(8.5)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)
            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)
            section.orientation = 1  # Landscape

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
        uni_run = uni_para.add_run(university_name.upper())
        uni_run.bold = True
        uni_run.font.size = Pt(18)
        uni_run.font.color.rgb = RGBColor(30, 64, 175)

        # Department
        dept_para = self._document.add_paragraph()
        dept_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dept_run = dept_para.add_run(f"Department of {department_name}")
        dept_run.bold = True
        dept_run.font.size = Pt(14)

        # Divider
        divider = self._document.add_paragraph()
        divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
        divider.add_run("═" * 80)

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
        stats = result.stats

        # Create a compact stats line
        stats_para = self._document.add_paragraph()
        stats_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        stats_text = (
            f"Total Desks: {stats.get('total_desks', 0)} | "
            f"Total Students: {stats.get('total_students', 0)} | "
            f"Cross-Stage Pairs: {stats.get('cross_stage_pairs', 0)} | "
            f"Empty Seats: {stats.get('empty_seats', 0)}"
        )
        stats_run = stats_para.add_run(stats_text)
        stats_run.font.size = Pt(10)
        stats_run.bold = True

        # Stage pattern
        if result.stage_order:
            pattern_para = self._document.add_paragraph()
            pattern_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pattern_run = pattern_para.add_run(
                f"Seating Pattern: {' → '.join(result.stage_order)} (repeating)"
            )
            pattern_run.font.size = Pt(10)
            pattern_run.italic = True

        # Warnings
        if result.warnings:
            warn_para = self._document.add_paragraph()
            warn_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            warn_run = warn_para.add_run(f"⚠ {' | '.join(result.warnings)}")
            warn_run.font.size = Pt(9)
            warn_run.font.color.rgb = RGBColor(180, 83, 9)

        self._document.add_paragraph()

    def _add_detailed_table(self, result: SeatingResult) -> None:
        """Add detailed seating table with all information."""
        # Section title
        title_para = self._document.add_paragraph()
        title_run = title_para.add_run("DETAILED SEATING LIST")
        title_run.bold = True
        title_run.font.size = Pt(12)

        # Create table
        table = self._document.add_table(rows=1, cols=7)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Header row
        headers = ["Desk #", "Column", "Row", "Seat A", "Student A", "Seat B", "Student B"]
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

            # Seat A number
            row.cells[3].text = str(desk.seat_a_number)
            self._style_cell(row.cells[3], align_center=True, bold=True)

            # Student A
            if desk.student_a:
                if desk.student_a.is_empty:
                    row.cells[4].text = "--- Empty ---"
                    self._style_cell(row.cells[4], bg_color=self.EMPTY_COLOR, align_center=True)
                else:
                    row.cells[4].text = f"{desk.student_a.name}\n({desk.student_a.stage})"
                    color = self.STAGE_COLORS.get(desk.student_a.stage, self.DEFAULT_COLOR)
                    self._style_cell(row.cells[4], bg_color=color)

            # Seat B number
            row.cells[5].text = str(desk.seat_b_number)
            self._style_cell(row.cells[5], align_center=True, bold=True)

            # Student B
            if desk.student_b:
                if desk.student_b.is_empty:
                    row.cells[6].text = "--- Empty ---"
                    self._style_cell(row.cells[6], bg_color=self.EMPTY_COLOR, align_center=True)
                else:
                    row.cells[6].text = f"{desk.student_b.name}\n({desk.student_b.stage})"
                    color = self.STAGE_COLORS.get(desk.student_b.stage, self.DEFAULT_COLOR)
                    self._style_cell(row.cells[6], bg_color=color)

        self._document.add_paragraph()

    def _add_visual_seating_map(self, result: SeatingResult) -> None:
        """Add visual seating map with desk cells."""
        # Page break
        self._document.add_page_break()

        # Title
        title_para = self._document.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run("VISUAL SEATING MAP")
        title_run.bold = True
        title_run.font.size = Pt(16)

        # Front label
        front_para = self._document.add_paragraph()
        front_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        front_run = front_para.add_run("▼ FRONT OF EXAMINATION HALL ▼")
        front_run.font.size = Pt(11)
        front_run.bold = True

        self._document.add_paragraph()

        # Column headers
        col_header_table = self._document.add_table(rows=1, cols=3)
        col_header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, col_name in enumerate(["RIGHT COLUMN", "MIDDLE COLUMN", "LEFT COLUMN"]):
            cell = col_header_table.rows[0].cells[i]
            cell.text = col_name
            self._style_cell(cell, bold=True, align_center=True, bg_color="374151", text_color="FFFFFF")
            cell.width = Cm(8)

        self._document.add_paragraph()

        # Group desks by row
        desks_by_row: Dict[int, Dict[str, Desk]] = {}
        max_row = 0

        for desk in result.desks:
            if desk.row not in desks_by_row:
                desks_by_row[desk.row] = {}
            desks_by_row[desk.row][desk.column.value] = desk
            max_row = max(max_row, desk.row)

        # Create visual map table - each cell is a DESK
        if max_row > 0:
            table = self._document.add_table(rows=max_row, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            # Set column widths
            for col in table.columns:
                for cell in col.cells:
                    cell.width = Cm(8)

            for row_num in range(1, max_row + 1):
                row_data = desks_by_row.get(row_num, {})
                table_row = table.rows[row_num - 1]

                for col_idx, col_name in enumerate(["Right", "Middle", "Left"]):
                    cell = table_row.cells[col_idx]
                    desk = row_data.get(col_name)

                    if desk:
                        self._format_desk_cell(cell, desk)
                    else:
                        cell.text = ""
                        self._style_cell(cell, bg_color="F9FAFB")

                    # Set row height
                    tr = table_row._tr
                    trPr = tr.get_or_add_trPr()
                    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="1200" w:hRule="atLeast"/>')
                    trPr.append(trHeight)

        self._document.add_paragraph()

        # Back label
        back_para = self._document.add_paragraph()
        back_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        back_run = back_para.add_run("▲ BACK OF EXAMINATION HALL ▲")
        back_run.font.size = Pt(11)
        back_run.bold = True

    def _format_desk_cell(self, cell, desk: Desk) -> None:
        """
        Format a desk cell with:
        - Student Name (Bold)
        - Stage Name (Smaller font)
        - Seat Number
        """
        cell.text = ""

        # Desk header
        p_header = cell.paragraphs[0]
        p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header_run = p_header.add_run(f"═══ Desk {desk.number} ═══")
        header_run.bold = True
        header_run.font.size = Pt(9)

        # Student A section
        p_a = cell.add_paragraph()
        p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_a.paragraph_format.space_before = Pt(4)
        p_a.paragraph_format.space_after = Pt(2)

        if desk.student_a:
            if desk.student_a.is_empty:
                # Empty seat
                empty_run = p_a.add_run(f"Seat {desk.seat_a_number}: --- Empty ---")
                empty_run.font.size = Pt(9)
                empty_run.italic = True
            else:
                # Seat number
                seat_run = p_a.add_run(f"Seat {desk.seat_a_number}: ")
                seat_run.font.size = Pt(8)
                seat_run.bold = True

                # Student name (Bold)
                name_a = desk.student_a.name
                if len(name_a) > 20:
                    name_a = name_a[:18] + "..."
                name_run = p_a.add_run(name_a)
                name_run.bold = True
                name_run.font.size = Pt(10)

                # Stage (smaller)
                stage_p = cell.add_paragraph()
                stage_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                stage_p.paragraph_format.space_before = Pt(0)
                stage_p.paragraph_format.space_after = Pt(4)
                stage_run = stage_p.add_run(f"({desk.student_a.stage})")
                stage_run.font.size = Pt(8)
                stage_run.italic = True

        # Separator
        sep_p = cell.add_paragraph()
        sep_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sep_p.paragraph_format.space_before = Pt(2)
        sep_p.paragraph_format.space_after = Pt(2)
        sep_run = sep_p.add_run("─────────")
        sep_run.font.size = Pt(8)

        # Student B section
        p_b = cell.add_paragraph()
        p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_b.paragraph_format.space_before = Pt(2)
        p_b.paragraph_format.space_after = Pt(2)

        if desk.student_b:
            if desk.student_b.is_empty:
                # Empty seat
                empty_run = p_b.add_run(f"Seat {desk.seat_b_number}: --- Empty ---")
                empty_run.font.size = Pt(9)
                empty_run.italic = True
            else:
                # Seat number
                seat_run = p_b.add_run(f"Seat {desk.seat_b_number}: ")
                seat_run.font.size = Pt(8)
                seat_run.bold = True

                # Student name (Bold)
                name_b = desk.student_b.name
                if len(name_b) > 20:
                    name_b = name_b[:18] + "..."
                name_run = p_b.add_run(name_b)
                name_run.bold = True
                name_run.font.size = Pt(10)

                # Stage (smaller)
                stage_p2 = cell.add_paragraph()
                stage_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                stage_p2.paragraph_format.space_before = Pt(0)
                stage_p2.paragraph_format.space_after = Pt(2)
                stage_run2 = stage_p2.add_run(f"({desk.student_b.stage})")
                stage_run2.font.size = Pt(8)
                stage_run2.italic = True

        # Set cell background based on cross-stage status
        if desk.is_cross_stage:
            self._set_cell_shading(cell, "D1FAE5")  # Green - good
        elif desk.has_real_students:
            if not desk.student_a.is_empty and not desk.student_b.is_empty:
                if desk.student_a.stage == desk.student_b.stage:
                    self._set_cell_shading(cell, "FEF3C7")  # Yellow - same stage warning
                else:
                    self._set_cell_shading(cell, "D1FAE5")  # Green
            else:
                self._set_cell_shading(cell, "E5E7EB")  # Gray - has empty
        else:
            self._set_cell_shading(cell, "E5E7EB")  # Gray

        # Vertical alignment
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _add_legend(self, result: SeatingResult) -> None:
        """Add color legend."""
        self._document.add_paragraph()

        legend_para = self._document.add_paragraph()
        legend_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        legend_run = legend_para.add_run("LEGEND: ")
        legend_run.bold = True
        legend_run.font.size = Pt(10)

        legend_text = (
            "🟢 Green = Cross-Stage (Good) | "
            "🟡 Yellow = Same Stage (Warning) | "
            "⬜ Gray = Empty Seat"
        )
        text_run = legend_para.add_run(legend_text)
        text_run.font.size = Pt(9)

        # Stage colors legend
        if result.stage_order:
            stage_para = self._document.add_paragraph()
            stage_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            stage_para.add_run("Stage Colors: ").bold = True
            colors_text = " | ".join([
                f"{stage}: {self.STAGE_COLORS.get(stage, 'Gray')}"
                for stage in result.stage_order
            ])
            stage_para.add_run(colors_text).font.size = Pt(9)

    def _add_footer(self) -> None:
        """Add footer with signatures and timestamp."""
        self._document.add_paragraph()

        # Signature table
        sig_table = self._document.add_table(rows=2, cols=3)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        labels = ["Prepared by:", "Verified by:", "Approved by:"]
        for i, label in enumerate(labels):
            sig_table.cell(0, i).text = label
            sig_table.cell(1, i).text = "\n_____________________\nName & Signature"

        for row in sig_table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(10)

        # Timestamp
        self._document.add_paragraph()
        time_para = self._document.add_paragraph()
        time_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        time_run = time_para.add_run(
            f"Document Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
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
                    try:
                        run.font.color.rgb = RGBColor.from_string(text_color)
                    except Exception:
                        pass

        if bg_color:
            self._set_cell_shading(cell, bg_color)

    def _set_cell_shading(self, cell, color: str) -> None:
        """Set background color for a table cell."""
        try:
            shading = parse_xml(
                f'<w:shd {nsdecls("w")} w:fill="{color}" w:val="clear"/>'
            )
            cell._tc.get_or_add_tcPr().append(shading)
        except Exception:
            pass  # Ignore shading errors
