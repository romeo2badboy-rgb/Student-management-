"""
Data Manager Module - Advanced Version v3.0
Handles multi-department student data with improved Arabic name extraction.
إدارة بيانات الطلاب مع دعم محسن للأسماء العربية
"""

from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field
import os
import re


# Import handlers
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import openpyxl
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False


@dataclass
class Stage:
    """Represents an academic stage with its students."""
    name: str
    students: List[str] = field(default_factory=list)

    def add_student(self, name: str) -> bool:
        """Add a student if not already exists."""
        cleaned = name.strip()
        if cleaned and cleaned not in self.students:
            self.students.append(cleaned)
            return True
        return False

    def remove_student(self, name: str) -> bool:
        """Remove a student from this stage."""
        if name in self.students:
            self.students.remove(name)
            return True
        return False

    def clear(self) -> None:
        """Clear all students."""
        self.students.clear()

    @property
    def count(self) -> int:
        """Get student count."""
        return len(self.students)


@dataclass
class Department:
    """Represents a university department with multiple stages."""
    name: str
    stages: Dict[str, Stage] = field(default_factory=dict)

    def __post_init__(self):
        """Initialize default stages if not provided."""
        if not self.stages:
            self.stages = {
                "1st Stage": Stage("1st Stage"),
                "2nd Stage": Stage("2nd Stage"),
                "3rd Stage": Stage("3rd Stage")
            }

    def get_stage(self, stage_name: str) -> Optional[Stage]:
        """Get a stage by name."""
        return self.stages.get(stage_name)

    def get_all_students(self) -> Dict[str, List[str]]:
        """Get all students organized by stage."""
        return {name: stage.students.copy() for name, stage in self.stages.items()}

    def get_total_students(self) -> int:
        """Get total student count across all stages."""
        return sum(stage.count for stage in self.stages.values())

    def get_stages_with_students(self) -> List[str]:
        """Get stages that have students."""
        return [name for name, stage in self.stages.items() if stage.count > 0]

    def clear_all(self) -> None:
        """Clear all students from all stages."""
        for stage in self.stages.values():
            stage.clear()


class DataManager:
    """
    Advanced Data Manager for University Exam Seating System.
    Supports multiple departments with multi-stage student management.
    """

    SUPPORTED_STAGES = ["1st Stage", "2nd Stage", "3rd Stage"]

    # كلمات يجب تجاهلها عند استيراد الأسماء
    SKIP_WORDS = [
        'name', 'student', 'no', 'number', '#', 'stage', 'الاسم', 'رقم',
        'ت', 'التسلسل', 'اسم', 'الطالب', 'ملاحظات', 'القسم', 'المرحلة',
        'اولى', 'ثانية', 'ثالثة', 'رابعة', 'first', 'second', 'third',
        'total', 'المجموع', 'العدد', 'count', 'page', 'صفحة'
    ]

    def __init__(self):
        """Initialize the data manager."""
        self._departments: Dict[str, Department] = {}
        self._university_name: str = "اسم المؤسسة"
        self._active_department: Optional[str] = None

    # ==================== University Settings ====================

    @property
    def university_name(self) -> str:
        """Get university name."""
        return self._university_name

    @university_name.setter
    def university_name(self, name: str) -> None:
        """Set university name."""
        self._university_name = name.strip() if name else "اسم المؤسسة"

    # ==================== Department Management ====================

    @property
    def departments(self) -> List[str]:
        """Get list of department names."""
        return list(self._departments.keys())

    @property
    def active_department(self) -> Optional[str]:
        """Get currently active department."""
        return self._active_department

    @active_department.setter
    def active_department(self, name: str) -> None:
        """Set active department."""
        if name in self._departments:
            self._active_department = name

    def create_department(self, name: str) -> Tuple[bool, str]:
        """Create a new department."""
        cleaned_name = name.strip()
        if not cleaned_name:
            return False, "اسم القسم لا يمكن أن يكون فارغاً"

        if cleaned_name in self._departments:
            return False, f"القسم '{cleaned_name}' موجود مسبقاً"

        self._departments[cleaned_name] = Department(cleaned_name)
        if self._active_department is None:
            self._active_department = cleaned_name

        return True, f"تم إنشاء القسم '{cleaned_name}' بنجاح"

    def delete_department(self, name: str) -> Tuple[bool, str]:
        """Delete a department."""
        if name not in self._departments:
            return False, f"القسم '{name}' غير موجود"

        del self._departments[name]

        if self._active_department == name:
            self._active_department = next(iter(self._departments), None)

        return True, f"تم حذف القسم '{name}' بنجاح"

    def get_department(self, name: str) -> Optional[Department]:
        """Get a department by name."""
        return self._departments.get(name)

    def get_active_department_obj(self) -> Optional[Department]:
        """Get the active department object."""
        if self._active_department:
            return self._departments.get(self._active_department)
        return None

    # ==================== Student Management ====================

    def get_students(self, department: str, stage: str) -> List[str]:
        """Get students for a specific department and stage."""
        dept = self._departments.get(department)
        if dept:
            stage_obj = dept.get_stage(stage)
            if stage_obj:
                return stage_obj.students.copy()
        return []

    def get_student_count(self, department: str, stage: str) -> int:
        """Get student count for a specific department and stage."""
        dept = self._departments.get(department)
        if dept:
            stage_obj = dept.get_stage(stage)
            if stage_obj:
                return stage_obj.count
        return 0

    def add_student(self, department: str, stage: str, name: str) -> bool:
        """Add a student to a department stage."""
        dept = self._departments.get(department)
        if dept:
            stage_obj = dept.get_stage(stage)
            if stage_obj:
                return stage_obj.add_student(name)
        return False

    def remove_student(self, department: str, stage: str, name: str) -> bool:
        """Remove a student from a department stage."""
        dept = self._departments.get(department)
        if dept:
            stage_obj = dept.get_stage(stage)
            if stage_obj:
                return stage_obj.remove_student(name)
        return False

    def clear_stage(self, department: str, stage: str) -> bool:
        """Clear all students from a specific stage."""
        dept = self._departments.get(department)
        if dept:
            stage_obj = dept.get_stage(stage)
            if stage_obj:
                stage_obj.clear()
                return True
        return False

    # ==================== جمع جميع الطلاب من جميع الأقسام ====================

    def get_all_students_all_departments(self) -> Dict[str, Dict[str, List[str]]]:
        """
        جمع جميع الطلاب من جميع الأقسام والمراحل
        Returns: {department_name: {stage_name: [students]}}
        """
        result = {}
        for dept_name, dept in self._departments.items():
            result[dept_name] = dept.get_all_students()
        return result

    def get_all_students_flat(self) -> List[Tuple[str, str, str]]:
        """
        جمع جميع الطلاب بشكل مسطح
        Returns: [(student_name, stage_name, department_name), ...]
        """
        students = []
        for dept_name, dept in self._departments.items():
            for stage_name, stage in dept.stages.items():
                for student in stage.students:
                    students.append((student, stage_name, dept_name))
        return students

    def get_total_students_all(self) -> int:
        """الحصول على إجمالي عدد الطلاب في جميع الأقسام"""
        total = 0
        for dept in self._departments.values():
            total += dept.get_total_students()
        return total

    # ==================== File Import ====================

    def import_from_file(
        self,
        file_path: str,
        department: str,
        stage: str
    ) -> Tuple[bool, str, int]:
        """Import students from a .docx or .xlsx file."""
        if department not in self._departments:
            return False, f"القسم '{department}' غير موجود", 0

        dept = self._departments[department]
        if stage not in dept.stages:
            return False, f"المرحلة غير صالحة: {stage}", 0

        if not os.path.exists(file_path):
            return False, f"الملف غير موجود: {file_path}", 0

        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            if file_ext == '.docx':
                return self._import_docx(file_path, dept, stage)
            elif file_ext in ['.xlsx', '.xls']:
                return self._import_xlsx(file_path, dept, stage)
            else:
                return False, f"صيغة ملف غير مدعومة: {file_ext}. استخدم .docx أو .xlsx", 0

        except PermissionError:
            return False, "تم رفض الإذن. يرجى إغلاق الملف إذا كان مفتوحاً", 0
        except Exception as e:
            return False, f"خطأ في استيراد الملف: {str(e)}", 0

    def _import_docx(
        self,
        file_path: str,
        department: Department,
        stage: str
    ) -> Tuple[bool, str, int]:
        """Import students from a Word document with improved name extraction."""
        if not DOCX_AVAILABLE:
            return False, "مكتبة python-docx غير مثبتة. شغل: pip install python-docx", 0

        try:
            doc = Document(file_path)
            stage_obj = department.stages[stage]
            imported_count = 0
            all_text_lines = []

            # جمع كل النصوص من الفقرات
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    all_text_lines.append(text)

            # جمع النصوص من الجداول
            for table in doc.tables:
                for row in table.rows:
                    row_texts = []
                    for cell in row.cells:
                        cell_text = cell.text.strip()
                        if cell_text:
                            row_texts.append(cell_text)
                    # معالجة كل صف
                    for text in row_texts:
                        all_text_lines.append(text)

            # استخراج الأسماء من جميع النصوص
            for line in all_text_lines:
                names = self._extract_arabic_names(line)
                for name in names:
                    if stage_obj.add_student(name):
                        imported_count += 1

            if imported_count == 0:
                return False, "لم يتم العثور على أسماء صالحة في المستند", 0

            return True, f"تم استيراد {imported_count} طالب إلى {stage}", imported_count

        except Exception as e:
            return False, f"خطأ في قراءة ملف .docx: {str(e)}", 0

    def _import_xlsx(
        self,
        file_path: str,
        department: Department,
        stage: str
    ) -> Tuple[bool, str, int]:
        """Import students from an Excel file with improved name extraction."""
        if not XLSX_AVAILABLE:
            return False, "مكتبة openpyxl غير مثبتة. شغل: pip install openpyxl", 0

        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            stage_obj = department.stages[stage]
            imported_count = 0

            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell_value in row:
                        if cell_value is not None:
                            names = self._extract_arabic_names(str(cell_value))
                            for name in names:
                                if stage_obj.add_student(name):
                                    imported_count += 1

            workbook.close()

            if imported_count == 0:
                return False, "لم يتم العثور على أسماء صالحة في ملف Excel", 0

            return True, f"تم استيراد {imported_count} طالب إلى {stage}", imported_count

        except Exception as e:
            return False, f"خطأ في قراءة ملف .xlsx: {str(e)}", 0

    def _extract_arabic_names(self, text: str) -> List[str]:
        """
        استخراج الأسماء العربية الثلاثية/الرباعية بشكل صحيح
        يتجاهل الأرقام والكلمات غير المفيدة
        """
        if not text:
            return []

        names = []

        # تنظيف النص
        text = text.strip()

        # إزالة الأرقام في بداية السطر (مثل: 1. أو 1- أو 1))
        text = re.sub(r'^[\d]+[\.\-\)\]\:\s]+', '', text)

        # تقسيم بالفواصل والأسطر الجديدة
        for delimiter in ['\n', '\r', ';', '،', '|']:
            text = text.replace(delimiter, ',')

        parts = text.split(',')

        for part in parts:
            cleaned = part.strip()

            # إزالة الأرقام في البداية مرة أخرى
            cleaned = re.sub(r'^[\d]+[\.\-\)\]\:\s]*', '', cleaned)
            cleaned = cleaned.strip()

            if not cleaned:
                continue

            # تجاهل إذا كان رقماً فقط
            if cleaned.isdigit():
                continue

            # تجاهل إذا كان قصيراً جداً (أقل من 4 أحرف)
            if len(cleaned) < 4:
                continue

            # تجاهل الكلمات المحجوزة
            if cleaned.lower() in self.SKIP_WORDS or cleaned in self.SKIP_WORDS:
                continue

            # التحقق من أن النص يحتوي على أحرف عربية أو إنجليزية
            has_arabic = bool(re.search(r'[\u0600-\u06FF]', cleaned))
            has_english = bool(re.search(r'[a-zA-Z]', cleaned))

            if not (has_arabic or has_english):
                continue

            # التحقق من أن الاسم يحتوي على كلمتين على الأقل (اسم ثنائي)
            words = cleaned.split()

            # تصفية الكلمات - إزالة الأرقام والكلمات القصيرة جداً
            valid_words = []
            for word in words:
                word = word.strip()
                # إزالة الأرقام من الكلمة
                word = re.sub(r'[\d]+', '', word)
                word = word.strip()

                if len(word) >= 2 and not word.isdigit():
                    valid_words.append(word)

            # يجب أن يكون الاسم ثنائياً على الأقل
            if len(valid_words) >= 2:
                full_name = ' '.join(valid_words)

                # التحقق النهائي من عدم وجود كلمات محجوزة
                is_valid = True
                for skip in self.SKIP_WORDS:
                    if skip in full_name.lower():
                        is_valid = False
                        break

                if is_valid and full_name not in names:
                    names.append(full_name)

        return names

    def _extract_names(self, text: str) -> List[str]:
        """Extract student names from text - legacy method."""
        return self._extract_arabic_names(text)

    # ==================== Validation ====================

    def validate_for_seating(self, department: str) -> Tuple[bool, str]:
        """Validate if a department is ready for seating generation."""
        dept = self._departments.get(department)
        if not dept:
            return False, f"القسم '{department}' غير موجود"

        stages_with_students = dept.get_stages_with_students()

        if len(stages_with_students) < 2:
            return False, "يجب وجود مرحلتين على الأقل مع طلاب لمنع الغش"

        total = dept.get_total_students()
        if total < 2:
            return False, "يجب وجود طالبين على الأقل"

        return True, "البيانات جاهزة للتوزيع"

    def validate_for_comprehensive_seating(self) -> Tuple[bool, str]:
        """التحقق من جاهزية البيانات للترتيب الشامل"""
        if len(self._departments) == 0:
            return False, "لا توجد أقسام. أنشئ قسماً واحداً على الأقل"

        total_students = self.get_total_students_all()
        if total_students < 2:
            return False, "يجب وجود طالبين على الأقل في جميع الأقسام"

        # التحقق من وجود مرحلتين مختلفتين على الأقل
        all_stages = set()
        for dept in self._departments.values():
            for stage_name, stage in dept.stages.items():
                if stage.count > 0:
                    all_stages.add(stage_name)

        if len(all_stages) < 2:
            return False, "يجب وجود مرحلتين مختلفتين على الأقل (في أي قسم) لمنع الغش"

        return True, f"جاهز! {total_students} طالب في {len(self._departments)} قسم"

    # ==================== Statistics ====================

    def get_department_summary(self, department: str) -> Dict:
        """Get summary statistics for a department."""
        dept = self._departments.get(department)
        if not dept:
            return {}

        summary = {
            "name": dept.name,
            "total_students": dept.get_total_students(),
            "stages": {}
        }

        for stage_name, stage in dept.stages.items():
            summary["stages"][stage_name] = stage.count

        return summary

    def get_global_summary(self) -> Dict:
        """Get summary of all departments."""
        return {
            "university": self._university_name,
            "total_departments": len(self._departments),
            "departments": [
                self.get_department_summary(name)
                for name in self._departments
            ]
        }

    # ==================== Data Export ====================

    def export_data(self) -> Dict:
        """Export all data as a dictionary (for saving/loading)."""
        return {
            "university_name": self._university_name,
            "active_department": self._active_department,
            "departments": {
                name: {
                    "stages": {
                        s_name: stage.students.copy()
                        for s_name, stage in dept.stages.items()
                    }
                }
                for name, dept in self._departments.items()
            }
        }

    def import_data(self, data: Dict) -> bool:
        """Import data from a dictionary."""
        try:
            self._university_name = data.get("university_name", "اسم المؤسسة")
            self._departments.clear()

            for dept_name, dept_data in data.get("departments", {}).items():
                self.create_department(dept_name)
                for stage_name, students in dept_data.get("stages", {}).items():
                    for student in students:
                        self.add_student(dept_name, stage_name, student)

            self._active_department = data.get("active_department")
            if self._active_department not in self._departments:
                self._active_department = next(iter(self._departments), None)

            return True
        except Exception:
            return False
