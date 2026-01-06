"""
Word Exporter Module v3.0 - Arabic Version
تصدير تقارير توزيع المقاعد بصيغة عربية

Features:
- دعم القاعات المتعددة
- عرض اسم المدير وعدد الطلاب
- تصدير الترتيب الشامل
- تنسيق القطاعات مع علامة الباب
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
import os

from seating_algorithm import SeatingResult, Desk, Column, Room


class WordExporter:
    """
    Arabic Word Document Exporter for Exam Seating.
    مُصدّر مستندات Word لتوزيع المقاعد للامتحانات

    Features:
    - رقم القاعة
    - القطاعات مع جداول المقاعد
    - علامة الباب
    - عدد الطلاب
    - توقيع مدير الاعدادية
    - دعم القاعات المتعددة
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
        "comprehensive_title": "توزيع شامل للامتحانات",
        "total_rooms": "عدد القاعات",
        "total_students_all": "إجمالي الطلاب",
        "page_of": "صفحة",
        "from": "من",
    }

    # Desks per sector: 3 columns × 4 rows = 12 (but door takes one, so 11 effective)
    DESKS_PER_SECTOR = 12

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
        Export seating arrangement to Word document.
        تصدير التوزيع إلى مستند Word

        للتوزيع الشامل (القاعات المتعددة) يتم تصدير كل قاعة في صفحة منفصلة
        """
        if not result:
            return False, "لا توجد بيانات للتصدير"

        # Check if comprehensive (multiple rooms)
        if result.rooms:
            return self._export_comprehensive(
                result=result,
                university_name=university_name,
                output_path=output_path,
                exam_title=exam_title,
                director_name=director_name
            )

        # Single room export
        if not result.desks:
            return False, "لا توجد بيانات للتصدير"

        if not output_path.lower().endswith('.docx'):
            output_path += '.docx'

        try:
            self._document = Document()
            self._setup_document()

            total_students = result.total_students
            desks_needed = result.total_desks
            sectors_needed = max(1, math.ceil(desks_needed / self.DESKS_PER_SECTOR))

            # Add room header
            self._add_room_header(room_number, total_students, director_name)

            # Add sectors
            desk_index = 0
            for sector_num in range(1, sectors_needed + 1):
                sector_desks = result.desks[desk_index:desk_index + self.DESKS_PER_SECTOR]
                desk_index += self.DESKS_PER_SECTOR

                self._add_sector(sector_num, sector_desks, is_first=(sector_num == 1))

                if sector_num < sectors_needed:
                    self._document.add_paragraph()

            # Add footer with student count and director
            self._add_footer(total_students, director_name)

            self._document.save(output_path)
            return True, f"تم التصدير بنجاح إلى {output_path}"

        except PermissionError:
            return False, f"تم رفض الإذن. يرجى إغلاق الملف: {output_path}"
        except Exception as e:
            return False, f"خطأ في التصدير: {str(e)}"

    def _export_comprehensive(
        self,
        result: SeatingResult,
        university_name: str,
        output_path: str,
        exam_title: str,
        director_name: str
    ) -> Tuple[bool, str]:
        """
        Export comprehensive seating (multiple rooms) to Word documents.
        تصدير الترتيب الشامل - كل قاعة في ملف منفصل أو صفحات منفصلة
        """
        if not result.rooms:
            return False, "لا توجد قاعات للتصدير"

        if not output_path.lower().endswith('.docx'):
            output_path += '.docx'

        try:
            # Create base directory path
            base_dir = os.path.dirname(output_path)
            base_name = os.path.splitext(os.path.basename(output_path))[0]

            exported_files = []
            total_rooms = len(result.rooms)
            total_students = result.total_students

            # Export each room to a separate file
            for room in result.rooms:
                room_file = os.path.join(
                    base_dir,
                    f"{base_name}_قاعة_{room.number}.docx"
                )

                self._document = Document()
                self._setup_document()

                room_students = room.total_students

                # Room header with info
                self._add_room_header(
                    room_number=room.number,
                    total_students=room_students,
                    director_name=director_name,
                    total_rooms=total_rooms,
                    total_students_all=total_students
                )

                # Add sectors for this room
                desks_needed = len(room.desks)
                sectors_needed = max(1, math.ceil(desks_needed / self.DESKS_PER_SECTOR))

                desk_index = 0
                for sector_num in range(1, sectors_needed + 1):
                    sector_desks = room.desks[desk_index:desk_index + self.DESKS_PER_SECTOR]
                    desk_index += self.DESKS_PER_SECTOR

                    self._add_sector(sector_num, sector_desks, is_first=(sector_num == 1))

                    if sector_num < sectors_needed:
                        self._document.add_paragraph()

                # Footer
                self._add_footer(room_students, director_name)

                self._document.save(room_file)
                exported_files.append(room_file)

            return True, f"تم تصدير {len(exported_files)} ملف بنجاح:\n" + "\n".join(exported_files)

        except PermissionError:
            return False, "تم رفض الإذن. يرجى إغلاق الملفات المفتوحة"
        except Exception as e:
            return False, f"خطأ في التصدير: {str(e)}"

    def _setup_document(self) -> None:
        """Configure document settings for Arabic RTL."""
        for section in self._document.sections:
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)
            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)

    def _set_rtl_paragraph(self, paragraph) -> None:
        """Set paragraph to RTL for Arabic."""
        try:
            pPr = paragraph._p.get_or_add_pPr()
            bidi = parse_xml(f'<w:bidi {nsdecls("w")} w:val="1"/>')
            pPr.append(bidi)
        except Exception:
            pass

    def _add_room_header(
        self,
        room_number: int,
        total_students: int,
        director_name: str = "",
        total_rooms: int = 1,
        total_students_all: int = 0
    ) -> None:
        """
        Add room header with information.
        إضافة رأس القاعة مع المعلومات
        """
        # Room number (centered, large)
        para = self._document.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(para)

        run = para.add_run(f"{self.LABELS['room_number']} ({room_number})")
        run.bold = True
        run.font.size = Pt(18)
        run.font.name = "Arial"

        # If comprehensive, show total rooms info
        if total_rooms > 1:
            info_para = self._document.add_paragraph()
            info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_rtl_paragraph(info_para)

            info_run = info_para.add_run(
                f"({self.LABELS['page_of']} {room_number} {self.LABELS['from']} {total_rooms})"
            )
            info_run.font.size = Pt(10)
            info_run.font.name = "Arial"
            info_run.italic = True

        # Separator line
        self._document.add_paragraph()

    def _add_sector(self, sector_num: int, desks: List[Desk], is_first: bool = False) -> None:
        """
        Add sector with seating table.
        إضافة قطاع مع جدول التوزيع
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

        # Set table dimensions
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell.width = Cm(4.5)

                tr = row._tr
                trPr = tr.get_or_add_trPr()
                try:
                    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="900" w:hRule="atLeast"/>')
                    trPr.append(trHeight)
                except Exception:
                    pass

        # Add door marker in first sector, first row, last column (right side)
        if is_first:
            door_cell = table.rows[0].cells[num_cols - 1]
            self._add_door_marker(door_cell)

        # Fill desks: Right to Left, Top to Bottom
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
                    self._make_empty_cell(cell)

        self._document.add_paragraph()

    def _add_door_marker(self, cell) -> None:
        """Add door marker with X pattern."""
        cell.text = ""

        p1 = cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p1)

        run1 = p1.add_run(self.LABELS['door'])
        run1.font.size = Pt(10)
        run1.font.name = "Arial"
        run1.bold = True

        # X pattern
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        x_run = p2.add_run("\\          /")
        x_run.font.size = Pt(12)

        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        x_run2 = p3.add_run("/          \\")
        x_run2.font.size = Pt(12)

        p4 = cell.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p4)
        run4 = p4.add_run(self.LABELS['door'])
        run4.font.size = Pt(10)
        run4.font.name = "Arial"
        run4.bold = True

        self._set_cell_shading(cell, "DCDCDC")
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _fill_desk_cell(self, cell, desk: Desk) -> None:
        """
        Fill cell with student names (no stage labels).
        تعبئة الخلية بأسماء الطلاب فقط
        """
        cell.text = ""

        # Student A
        p_a = cell.paragraphs[0]
        p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p_a)

        if desk.student_a and not desk.student_a.is_empty:
            name_run = p_a.add_run(desk.student_a.name)
            name_run.bold = True
            name_run.font.size = Pt(10)
            name_run.font.name = "Arial"

        # Separator
        sep_p = cell.add_paragraph()
        sep_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sep_run = sep_p.add_run("─────────")
        sep_run.font.size = Pt(8)

        # Student B
        p_b = cell.add_paragraph()
        p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl_paragraph(p_b)

        if desk.student_b and not desk.student_b.is_empty:
            name_run_b = p_b.add_run(desk.student_b.name)
            name_run_b.bold = True
            name_run_b.font.size = Pt(10)
            name_run_b.font.name = "Arial"

        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _make_empty_cell(self, cell) -> None:
        """Make empty cell."""
        cell.text = ""
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _add_footer(self, total_students: int, director_name: str) -> None:
        """
        Add footer with student count and director signature.
        إضافة الذيل مع عدد الطلاب وتوقيع المدير
        """
        self._document.add_paragraph()

        # Student count (right aligned)
        count_para = self._document.add_paragraph()
        count_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl_paragraph(count_para)

        count_run = count_para.add_run(f"{self.LABELS['student_count']} / {total_students}")
        count_run.bold = True
        count_run.font.size = Pt(13)
        count_run.font.name = "Arial"

        self._document.add_paragraph()
        self._document.add_paragraph()

        # Director signature
        dir_para = self._document.add_paragraph()
        dir_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl_paragraph(dir_para)

        dir_run = dir_para.add_run(self.LABELS['director'])
        dir_run.font.size = Pt(12)
        dir_run.font.name = "Arial"

        # Director name (if provided)
        if director_name:
            name_para = self._document.add_paragraph()
            name_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            self._set_rtl_paragraph(name_para)

            name_run = name_para.add_run(director_name)
            name_run.bold = True
            name_run.font.size = Pt(13)
            name_run.font.name = "Arial"

    def _get_arabic_stage(self, stage: str) -> str:
        """Convert stage name to Arabic."""
        stage_map = {
            "1st Stage": self.LABELS['stage_1'],
            "2nd Stage": self.LABELS['stage_2'],
            "3rd Stage": self.LABELS['stage_3'],
        }
        return stage_map.get(stage, stage)

    def _add_student_count(self, total_students: int) -> None:
        """Add student count (legacy method)."""
        para = self._document.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl_paragraph(para)

        run = para.add_run(f"{self.LABELS['student_count']} / {total_students}")
        run.font.size = Pt(12)
        run.font.name = "Arial"

        self._document.add_paragraph()

    def _add_director_signature(self, director_name: str = "") -> None:
        """Add director signature (legacy method)."""
        title_para = self._document.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl_paragraph(title_para)

        title_run = title_para.add_run(self.LABELS['director'])
        title_run.font.size = Pt(12)
        title_run.font.name = "Arial"

        if director_name:
            name_para = self._document.add_paragraph()
            name_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            self._set_rtl_paragraph(name_para)

            name_run = name_para.add_run(director_name)
            name_run.font.size = Pt(12)
            name_run.font.name = "Arial"

    def _set_cell_shading(self, cell, color: str) -> None:
        """Set cell background color."""
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
