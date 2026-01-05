"""
Data Manager Module - Advanced Version
Handles multi-department student data with support for .docx and .xlsx imports.
"""

from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field
import os

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

    def __init__(self):
        """Initialize the data manager."""
        self._departments: Dict[str, Department] = {}
        self._university_name: str = "University Name"
        self._active_department: Optional[str] = None

    # ==================== University Settings ====================

    @property
    def university_name(self) -> str:
        """Get university name."""
        return self._university_name

    @university_name.setter
    def university_name(self, name: str) -> None:
        """Set university name."""
        self._university_name = name.strip() if name else "University Name"

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
        """
        Create a new department.

        Args:
            name: Department name

        Returns:
            Tuple of (success, message)
        """
        cleaned_name = name.strip()
        if not cleaned_name:
            return False, "Department name cannot be empty."

        if cleaned_name in self._departments:
            return False, f"Department '{cleaned_name}' already exists."

        self._departments[cleaned_name] = Department(cleaned_name)
        if self._active_department is None:
            self._active_department = cleaned_name

        return True, f"Department '{cleaned_name}' created successfully."

    def delete_department(self, name: str) -> Tuple[bool, str]:
        """
        Delete a department.

        Args:
            name: Department name to delete

        Returns:
            Tuple of (success, message)
        """
        if name not in self._departments:
            return False, f"Department '{name}' not found."

        del self._departments[name]

        # Update active department if needed
        if self._active_department == name:
            self._active_department = next(iter(self._departments), None)

        return True, f"Department '{name}' deleted successfully."

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

    # ==================== File Import ====================

    def import_from_file(
        self,
        file_path: str,
        department: str,
        stage: str
    ) -> Tuple[bool, str, int]:
        """
        Import students from a .docx or .xlsx file.

        Args:
            file_path: Path to the file
            department: Target department name
            stage: Target stage name

        Returns:
            Tuple of (success, message, count)
        """
        # Validate inputs
        if department not in self._departments:
            return False, f"Department '{department}' not found.", 0

        dept = self._departments[department]
        if stage not in dept.stages:
            return False, f"Invalid stage: {stage}", 0

        if not os.path.exists(file_path):
            return False, f"File not found: {file_path}", 0

        # Determine file type and import
        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            if file_ext == '.docx':
                return self._import_docx(file_path, dept, stage)
            elif file_ext in ['.xlsx', '.xls']:
                return self._import_xlsx(file_path, dept, stage)
            else:
                return False, f"Unsupported file format: {file_ext}. Use .docx or .xlsx", 0

        except PermissionError:
            return False, "Permission denied. Please close the file if it's open.", 0
        except Exception as e:
            return False, f"Error importing file: {str(e)}", 0

    def _import_docx(
        self,
        file_path: str,
        department: Department,
        stage: str
    ) -> Tuple[bool, str, int]:
        """Import students from a Word document."""
        if not DOCX_AVAILABLE:
            return False, "python-docx library not installed. Run: pip install python-docx", 0

        try:
            doc = Document(file_path)
            stage_obj = department.stages[stage]
            imported_count = 0

            # Extract from paragraphs
            for para in doc.paragraphs:
                names = self._extract_names(para.text)
                for name in names:
                    if stage_obj.add_student(name):
                        imported_count += 1

            # Extract from tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        names = self._extract_names(cell.text)
                        for name in names:
                            if stage_obj.add_student(name):
                                imported_count += 1

            if imported_count == 0:
                return False, "No valid student names found in the document.", 0

            return True, f"Successfully imported {imported_count} students to {stage}.", imported_count

        except Exception as e:
            return False, f"Error reading .docx file: {str(e)}", 0

    def _import_xlsx(
        self,
        file_path: str,
        department: Department,
        stage: str
    ) -> Tuple[bool, str, int]:
        """Import students from an Excel file."""
        if not XLSX_AVAILABLE:
            return False, "openpyxl library not installed. Run: pip install openpyxl", 0

        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            stage_obj = department.stages[stage]
            imported_count = 0

            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell_value in row:
                        if cell_value is not None:
                            names = self._extract_names(str(cell_value))
                            for name in names:
                                if stage_obj.add_student(name):
                                    imported_count += 1

            workbook.close()

            if imported_count == 0:
                return False, "No valid student names found in the Excel file.", 0

            return True, f"Successfully imported {imported_count} students to {stage}.", imported_count

        except Exception as e:
            return False, f"Error reading .xlsx file: {str(e)}", 0

    def _extract_names(self, text: str) -> List[str]:
        """Extract student names from text."""
        if not text:
            return []

        # Split by common delimiters
        names = []
        for delimiter in ['\n', ',', ';', '\t']:
            text = text.replace(delimiter, '|')

        parts = text.split('|')
        for part in parts:
            cleaned = part.strip()
            # Filter out numbers-only entries and very short strings
            if cleaned and len(cleaned) >= 2 and not cleaned.isdigit():
                # Skip common header words
                skip_words = ['name', 'student', 'no', 'number', '#', 'stage', 'الاسم', 'رقم']
                if cleaned.lower() not in skip_words:
                    names.append(cleaned)

        return names

    # ==================== Validation ====================

    def validate_for_seating(self, department: str) -> Tuple[bool, str]:
        """
        Validate if a department is ready for seating generation.

        Args:
            department: Department name

        Returns:
            Tuple of (is_valid, message)
        """
        dept = self._departments.get(department)
        if not dept:
            return False, f"Department '{department}' not found."

        stages_with_students = dept.get_stages_with_students()

        if len(stages_with_students) < 2:
            return False, "At least 2 stages must have students for anti-cheating seating."

        total = dept.get_total_students()
        if total < 2:
            return False, "At least 2 students are required."

        return True, "Data is valid for seating generation."

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
            self._university_name = data.get("university_name", "University Name")
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
