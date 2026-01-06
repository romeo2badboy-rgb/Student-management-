"""
Word Exporter Module - Arabic Version
Creates exam seating reports in Arabic format with sectors layout.
تصدير تقارير توزيع المقاعد بصيغة عربية
"""

from typing import List, Optional, Dict, Tuple
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
import math

from seating_algorithm import SeatingResult, Desk, Column


class WordExporter:
    """
    Arabic Word Document Exporter for Exam Seating.
    مُصدّر مستندات Word للتوزيع المقاعد للامتحانات

    Creates reports with:
    - Room number header (رقم القاعة)
    - Sectors with desk tables (قطاع)
    - Door marker (الباب)
    - Student count (عدد الطلاب)
    - Director signature (مدير الاعدادية)
    """

    # Arabic Labels
    LABELS = {
        "room_number": "رقم القاعة",
        "sector": "قطاع",
        "door": "الباب",
        "student_count": "عدد الطلاب",
        "director": "مدير الاعدادية",
        "empty": "فارغ",
        "stage_1": "المرحلة الاولى",
        "stage_2": "المرحلة الثانية",
        "stage_3": "المرحلة الثالثة",
    }

    # Desks per sector
    DESKS_PER_SECTOR = 12  # 3 columns × 4 rows

    def __init__(self):
        """Initialize the exporter."""
        self._document: Optional[Document] = None

    def export(
        self,
        result: SeatingResult,
        university_name: str,
        department_name: str,
        output_path: str,
        exam_title: str = "توزيع مقاعد الامتحان",
        exam_date: Optional[str] = None,
        include_map: bool = True,
        director_name: str = "",
        room_number: int = 1
    ) -> Tuple[bool, str]:
        """
        Export seating arrangement to a Word document in Arabic format.

        Args:
            result: SeatingResult from the algorithm
            university_name: Name of the university
            department_name: Name of the department
            output_path: Path for the output file
            exam_title: Title for the document
            exam_date: Date of the exam (optional)
            include_map: Whether to include visual seating map
            director_name: Name of the director
            room_number: Room/Hall number

        Returns:
            Tuple of (success, message)
        """
        if not result or not result.desks:
            return False, "لا توجد بيانات للتصدير. No seating data to export."

        if not output_path.lower().endswith('.docx'):
            output_path += '.docx'

        try:
            self._document = Document()
            self._setup_document()

            # Calculate sectors needed
            total_students = result.total_students
            desks_needed = result.total_desks
            sectors_needed = max(1, math.ceil(desks_needed / self.DESKS_PER_SECTOR))

            # Add room number header
            self._add_room_header(room_number)

            # Add sectors with seating tables
            desk_index = 0
            for sector_num in range(1, sectors_needed + 1):
                # Get desks for this sector
                sector_desks = result.desks[desk_index:desk_index + self.DESKS_PER_SECTOR]
                desk_index += self.DESKS_PER_SECTOR

                self._add_sector(sector_num, sector_desks, is_first=(sector_num == 1))

                # Add space between sectors
                if sector_num < sectors_needed:
                    self._document.add_paragraph()

            # Add student count
            self._add_student_count(total_students)

            # Add director signature
            self._add_director_signature(director_name)

            self._document.save(output_path)
            return True, f"تم التصدير بنجاح إلى {output_path}"

        except PermissionError:
            return False, f"تم رفض الإذن. يرجى إغلاق الملف: {output_path}"
        except Exception as e:
            return False, f"خطأ في التصدير: {str(e)}"

    def _setup_document(self) -> None:
        """Configure document settings for Arabic RTL."""
        for section in self._document.sections:
            # Portrait orientation for Arabic format
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)
            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)

    def _set_rtl_paragraph(self, paragraph) -> None:
        """Set paragraph to RTL (Right-to-Left) for Arabic."""
        try:
            pPr = paragraph._p.get_or_add_pPr()
            bidi = parse_xml(f'<w:bidi {nsdecls("w")} w:val="1"/>')
            pPr.append(bidi)
        except Exception:
            pass

    def _add_room_header(self, room_number: int) -> None:
        """Add room number header - رقم القاعة"""
        para = self._document.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(para)

        run = para.add_run(f"{self.LABELS['room_number']} ({room_number})")
        run.bold = True
        run.font.size = Pt(16)
        run.font.name = "Arial"

        self._document.add_paragraph()

    def _add_sector(self, sector_num: int, desks: List[Desk], is_first: bool = False) -> None:
        """
        Add a sector with seating table.
        قطاع مع جدول التوزيع
        """
        # Sector title
        title_para = self._document.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(title_para)

        title_run = title_para.add_run(f"{self.LABELS['sector']} ( {sector_num} )")
        title_run.bold = True
        title_run.font.size = Pt(14)
        title_run.font.name = "Arial"

        # Create table: 3 columns × 4 rows
        num_rows = 4
        num_cols = 3

        table = self._document.add_table(rows=num_rows, cols=num_cols)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Set table width and cell properties
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell.width = Cm(4.5)

                # Set minimum row height
                tr = row._tr
                trPr = tr.get_or_add_trPr()
                try:
                    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="900" w:hRule="atLeast"/>')
                    trPr.append(trHeight)
                except Exception:
                    pass

        # Fill the first row, last cell with door marker (for first sector)
        if is_first:
            door_cell = table.rows[0].cells[num_cols - 1]
            self._add_door_marker(door_cell)

        # Fill desks into table cells
        # Layout: Right to Left, Top to Bottom (Arabic reading order)
        desk_idx = 0
        for row_idx in range(num_rows):
            for col_idx in range(num_cols - 1, -1, -1):  # Right to Left
                cell = table.rows[row_idx].cells[col_idx]

                # Skip door cell
                if is_first and row_idx == 0 and col_idx == num_cols - 1:
                    continue

                if desk_idx < len(desks):
                    desk = desks[desk_idx]
                    self._fill_desk_cell(cell, desk)
                    desk_idx += 1
                else:
                    # Empty cell
                    self._make_empty_cell(cell)

        self._document.add_paragraph()

    def _add_door_marker(self, cell) -> None:
        """Add door marker with X pattern - الباب"""
        cell.text = ""

        # Top line with "الباب"
        p1 = cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p1)

        # Draw X pattern
        door_text = f"{self.LABELS['door']}"
        run1 = p1.add_run(door_text)
        run1.font.size = Pt(10)
        run1.font.name = "Arial"

        # Add diagonal lines representation
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        x_run = p2.add_run("╲      ╱")
        x_run.font.size = Pt(12)

        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        x_run2 = p3.add_run("╱      ╲")
        x_run2.font.size = Pt(12)

        # Add second "الباب" at bottom
        p4 = cell.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p4)
        run4 = p4.add_run(self.LABELS['door'])
        run4.font.size = Pt(10)
        run4.font.name = "Arial"

        # Set cell background to light gray
        self._set_cell_shading(cell, "E5E7EB")
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _fill_desk_cell(self, cell, desk: Desk) -> None:
        """
        Fill a cell with student information (name only, no stage).
        تعبئة خلية باسم الطالب فقط
        """
        cell.text = ""

        # Student A (Top of desk) - Name only
        p_a = cell.paragraphs[0]
        p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p_a)

        if desk.student_a and not desk.student_a.is_empty:
            name_run = p_a.add_run(desk.student_a.name)
            name_run.bold = True
            name_run.font.size = Pt(11)
            name_run.font.name = "Arial"
        else:
            # Empty cell placeholder
            pass

        # Separator line
        sep_p = cell.add_paragraph()
        sep_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sep_run = sep_p.add_run("─────────")
        sep_run.font.size = Pt(8)

        # Student B (Bottom of desk) - Name only
        p_b = cell.add_paragraph()
        p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p_b)

        if desk.student_b and not desk.student_b.is_empty:
            name_run_b = p_b.add_run(desk.student_b.name)
            name_run_b.bold = True
            name_run_b.font.size = Pt(11)
            name_run_b.font.name = "Arial"
        else:
            # Empty cell placeholder
            pass

        # Set cell vertical alignment
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _make_empty_cell(self, cell) -> None:
        """Make an empty cell for no desk."""
        cell.text = ""
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _get_arabic_stage(self, stage: str) -> str:
        """Convert stage name to Arabic."""
        stage_map = {
            "1st Stage": self.LABELS['stage_1'],
            "2nd Stage": self.LABELS['stage_2'],
            "3rd Stage": self.LABELS['stage_3'],
        }
        return stage_map.get(stage, stage)

    def _add_student_count(self, total_students: int) -> None:
        """Add student count - عدد الطلاب"""
        para = self._document.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl_paragraph(para)

        run = para.add_run(f"{self.LABELS['student_count']} / {total_students}")
        run.font.size = Pt(12)
        run.font.name = "Arial"

        self._document.add_paragraph()

    def _add_director_signature(self, director_name: str = "") -> None:
        """Add director signature section - مدير الاعدادية"""
        # Director title
        title_para = self._document.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl_paragraph(title_para)

        title_run = title_para.add_run(self.LABELS['director'])
        title_run.font.size = Pt(12)
        title_run.font.name = "Arial"

        # Director name
        if director_name:
            name_para = self._document.add_paragraph()
            name_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            self._set_rtl_paragraph(name_para)

            name_run = name_para.add_run(director_name)
            name_run.font.size = Pt(12)
            name_run.font.name = "Arial"

    def _set_cell_shading(self, cell, color: str) -> None:
        """Set background color for a table cell."""
        try:
            shading = parse_xml(
                f'<w:shd {nsdecls("w")} w:fill="{color}" w:val="clear"/>'
            )
            cell._tc.get_or_add_tcPr().append(shading)
        except Exception:
            pass

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
