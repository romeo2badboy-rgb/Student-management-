"""
Data Manager Module
Handles student data storage and management by academic stages.
"""

from typing import Dict, List, Optional
from docx import Document
import os


class DataManager:
    """
    Manages student data pools organized by academic stages.
    Provides methods for importing, storing, and retrieving student information.
    """

    def __init__(self):
        """Initialize the data manager with empty student pools."""
        self._student_pools: Dict[str, List[str]] = {
            "1st Stage": [],
            "2nd Stage": [],
            "3rd Stage": []
        }
        self._department_name: str = ""

    @property
    def department_name(self) -> str:
        """Get the department name."""
        return self._department_name

    @department_name.setter
    def department_name(self, name: str) -> None:
        """Set the department name."""
        self._department_name = name.strip()

    @property
    def stages(self) -> List[str]:
        """Get list of available stages."""
        return list(self._student_pools.keys())

    def get_students(self, stage: str) -> List[str]:
        """
        Get students for a specific stage.

        Args:
            stage: The academic stage (e.g., "1st Stage")

        Returns:
            List of student names for the specified stage
        """
        return self._student_pools.get(stage, []).copy()

    def get_all_students(self) -> Dict[str, List[str]]:
        """
        Get all students organized by stage.

        Returns:
            Dictionary with stages as keys and student lists as values
        """
        return {stage: students.copy() for stage, students in self._student_pools.items()}

    def get_student_count(self, stage: str) -> int:
        """
        Get the number of students in a specific stage.

        Args:
            stage: The academic stage

        Returns:
            Number of students in the stage
        """
        return len(self._student_pools.get(stage, []))

    def get_total_students(self) -> int:
        """
        Get total number of students across all stages.

        Returns:
            Total student count
        """
        return sum(len(students) for students in self._student_pools.values())

    def add_student(self, stage: str, name: str) -> bool:
        """
        Add a single student to a stage.

        Args:
            stage: The academic stage
            name: Student name

        Returns:
            True if added successfully, False otherwise
        """
        if stage in self._student_pools:
            cleaned_name = name.strip()
            if cleaned_name and cleaned_name not in self._student_pools[stage]:
                self._student_pools[stage].append(cleaned_name)
                return True
        return False

    def remove_student(self, stage: str, name: str) -> bool:
        """
        Remove a student from a stage.

        Args:
            stage: The academic stage
            name: Student name

        Returns:
            True if removed successfully, False otherwise
        """
        if stage in self._student_pools and name in self._student_pools[stage]:
            self._student_pools[stage].remove(name)
            return True
        return False

    def clear_stage(self, stage: str) -> bool:
        """
        Clear all students from a specific stage.

        Args:
            stage: The academic stage to clear

        Returns:
            True if cleared successfully, False otherwise
        """
        if stage in self._student_pools:
            self._student_pools[stage] = []
            return True
        return False

    def clear_all(self) -> None:
        """Clear all student data from all stages."""
        for stage in self._student_pools:
            self._student_pools[stage] = []
        self._department_name = ""

    def import_from_docx(self, file_path: str, stage: str) -> tuple[bool, str, int]:
        """
        Import student names from a Word document.

        Expects each student name on a separate line or paragraph.

        Args:
            file_path: Path to the .docx file
            stage: The academic stage to import students into

        Returns:
            Tuple of (success: bool, message: str, count: int)
        """
        if stage not in self._student_pools:
            return False, f"Invalid stage: {stage}", 0

        if not os.path.exists(file_path):
            return False, f"File not found: {file_path}", 0

        if not file_path.lower().endswith('.docx'):
            return False, "Invalid file format. Please select a .docx file.", 0

        try:
            doc = Document(file_path)
            imported_count = 0

            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if text:
                    # Handle multiple names per line (comma or newline separated)
                    names = [n.strip() for n in text.replace('\n', ',').split(',')]
                    for name in names:
                        if name and self.add_student(stage, name):
                            imported_count += 1

            # Also check tables in the document
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            names = [n.strip() for n in text.replace('\n', ',').split(',')]
                            for name in names:
                                if name and self.add_student(stage, name):
                                    imported_count += 1

            if imported_count == 0:
                return False, "No valid student names found in the document.", 0

            return True, f"Successfully imported {imported_count} students to {stage}.", imported_count

        except Exception as e:
            return False, f"Error reading file: {str(e)}", 0

    def get_stages_with_students(self) -> List[str]:
        """
        Get list of stages that have at least one student.

        Returns:
            List of stage names with students
        """
        return [stage for stage, students in self._student_pools.items() if students]

    def validate_for_seating(self) -> tuple[bool, str]:
        """
        Validate if current data is suitable for seating generation.

        Returns:
            Tuple of (is_valid: bool, message: str)
        """
        stages_with_students = self.get_stages_with_students()

        if not self._department_name:
            return False, "Please enter a department name."

        if len(stages_with_students) < 2:
            return False, "At least 2 stages must have students for mixed seating."

        total = self.get_total_students()
        if total < 2:
            return False, "At least 2 students are required to generate seating."

        return True, "Data is valid for seating generation."

    def get_summary(self) -> str:
        """
        Get a summary of current data.

        Returns:
            Formatted string with data summary
        """
        lines = [f"Department: {self._department_name or 'Not set'}", ""]
        for stage in self._student_pools:
            count = len(self._student_pools[stage])
            lines.append(f"{stage}: {count} students")
        lines.append(f"\nTotal: {self.get_total_students()} students")
        return "\n".join(lines)
