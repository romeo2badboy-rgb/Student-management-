"""
Advanced Seating Algorithm Module
Implements the Anti-Cheating Zigzag Seating Pattern.
"""

import random
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
from enum import Enum


class Column(Enum):
    """Exam hall column positions."""
    RIGHT = "Right"
    MIDDLE = "Middle"
    LEFT = "Left"


@dataclass
class Student:
    """Represents a student with their stage information."""
    name: str
    stage: str

    def __str__(self) -> str:
        return f"{self.name} ({self.stage})"


@dataclass
class Desk:
    """Represents a desk with two seats."""
    number: int
    column: Column
    row: int
    student_a: Optional[Student] = None
    student_b: Optional[Student] = None

    @property
    def is_full(self) -> bool:
        """Check if desk has both students."""
        return self.student_a is not None and self.student_b is not None

    @property
    def is_empty(self) -> bool:
        """Check if desk is empty."""
        return self.student_a is None and self.student_b is None

    @property
    def is_cross_stage(self) -> bool:
        """Check if desk has students from different stages."""
        if not self.is_full:
            return False
        return self.student_a.stage != self.student_b.stage

    @property
    def student_count(self) -> int:
        """Get number of students at this desk."""
        count = 0
        if self.student_a:
            count += 1
        if self.student_b:
            count += 1
        return count


@dataclass
class SeatingResult:
    """Contains the complete seating arrangement result."""
    desks: List[Desk] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    stats: Dict = field(default_factory=dict)

    @property
    def total_desks(self) -> int:
        return len(self.desks)

    @property
    def total_students(self) -> int:
        return sum(desk.student_count for desk in self.desks)

    @property
    def cross_stage_count(self) -> int:
        return sum(1 for desk in self.desks if desk.is_cross_stage)

    @property
    def same_stage_count(self) -> int:
        return sum(1 for desk in self.desks if desk.is_full and not desk.is_cross_stage)


class ZigzagSeatingAlgorithm:
    """
    Advanced Anti-Cheating Seating Algorithm.

    Implements a zigzag pattern where:
    - Students from different stages sit at the same desk
    - Students are randomized within their stages
    - Desks are arranged in 3 columns (Right, Middle, Left)
    - Handles unbalanced stage counts gracefully
    """

    COLUMNS = [Column.RIGHT, Column.MIDDLE, Column.LEFT]
    DESKS_PER_ROW = 3  # One desk per column

    def __init__(self):
        """Initialize the algorithm."""
        self._result: Optional[SeatingResult] = None

    @property
    def result(self) -> Optional[SeatingResult]:
        """Get the last seating result."""
        return self._result

    def generate(self, student_pools: Dict[str, List[str]]) -> Tuple[bool, str, SeatingResult]:
        """
        Generate anti-cheating seating arrangement.

        Args:
            student_pools: Dict mapping stage names to student name lists

        Returns:
            Tuple of (success, message, result)
        """
        self._result = SeatingResult()

        # Validate input
        valid_pools = {k: v for k, v in student_pools.items() if v}
        if len(valid_pools) < 2:
            return False, "At least 2 stages with students required.", self._result

        # Create Student objects and shuffle within stages
        all_students: List[Student] = []
        stage_students: Dict[str, List[Student]] = {}

        for stage, names in valid_pools.items():
            students = [Student(name, stage) for name in names]
            random.shuffle(students)  # Randomize within stage
            stage_students[stage] = students
            all_students.extend(students)

        if len(all_students) < 2:
            return False, "At least 2 students required.", self._result

        # Generate zigzag seating
        try:
            desks = self._create_zigzag_seating(stage_students)
            self._result.desks = desks
            self._calculate_stats()

            # Generate warnings
            if self._result.same_stage_count > 0:
                self._result.warnings.append(
                    f"{self._result.same_stage_count} desk(s) have students from the same stage "
                    "due to unbalanced numbers."
                )

            message = (
                f"Generated {self._result.total_desks} desks for "
                f"{self._result.total_students} students. "
                f"Cross-stage pairs: {self._result.cross_stage_count}"
            )

            return True, message, self._result

        except Exception as e:
            return False, f"Algorithm error: {str(e)}", self._result

    def _create_zigzag_seating(self, stage_students: Dict[str, List[Student]]) -> List[Desk]:
        """
        Create zigzag seating pattern ensuring cross-stage pairing.

        The algorithm:
        1. Sort stages by student count (descending)
        2. Alternate between stages when pairing
        3. Use zigzag pattern across columns
        4. Handle remainders by pairing from same stage if necessary
        """
        desks: List[Desk] = []

        # Create queues from each stage
        queues: Dict[str, List[Student]] = {
            stage: students.copy()
            for stage, students in stage_students.items()
        }

        # Get list of stages sorted by count (largest first)
        stages = sorted(queues.keys(), key=lambda s: len(queues[s]), reverse=True)

        desk_number = 1
        row = 1
        col_index = 0

        # Phase 1: Create cross-stage pairs using zigzag
        while self._has_multiple_active_stages(queues):
            # Get two different stages with students
            stage_a, stage_b = self._get_two_different_stages(queues, stages)

            if stage_a is None or stage_b is None:
                break

            student_a = queues[stage_a].pop(0)
            student_b = queues[stage_b].pop(0)

            # Create desk with zigzag column assignment
            column = self.COLUMNS[col_index % 3]

            desk = Desk(
                number=desk_number,
                column=column,
                row=row,
                student_a=student_a,
                student_b=student_b
            )
            desks.append(desk)

            desk_number += 1
            col_index += 1

            # Move to next row after filling all columns
            if col_index % 3 == 0:
                row += 1

            # Re-sort stages to balance distribution
            stages = sorted(
                [s for s in stages if queues[s]],
                key=lambda s: len(queues[s]),
                reverse=True
            )

        # Phase 2: Handle remaining students (same-stage pairs if necessary)
        remaining: List[Student] = []
        for stage in queues:
            remaining.extend(queues[stage])

        # Pair remaining students
        while len(remaining) >= 2:
            student_a = remaining.pop(0)
            student_b = remaining.pop(0)

            column = self.COLUMNS[col_index % 3]

            desk = Desk(
                number=desk_number,
                column=column,
                row=row,
                student_a=student_a,
                student_b=student_b
            )
            desks.append(desk)

            desk_number += 1
            col_index += 1

            if col_index % 3 == 0:
                row += 1

        # Phase 3: Handle single remaining student
        if remaining:
            student = remaining.pop(0)
            column = self.COLUMNS[col_index % 3]

            desk = Desk(
                number=desk_number,
                column=column,
                row=row,
                student_a=student,
                student_b=None
            )
            desks.append(desk)

        return desks

    def _has_multiple_active_stages(self, queues: Dict[str, List[Student]]) -> bool:
        """Check if there are at least 2 stages with students."""
        active_count = sum(1 for students in queues.values() if students)
        return active_count >= 2

    def _get_two_different_stages(
        self,
        queues: Dict[str, List[Student]],
        stages: List[str]
    ) -> Tuple[Optional[str], Optional[str]]:
        """Get two different stages that have students."""
        active_stages = [s for s in stages if queues[s]]

        if len(active_stages) < 2:
            return None, None

        # Return first two active stages (already sorted by count)
        return active_stages[0], active_stages[1]

    def _calculate_stats(self) -> None:
        """Calculate statistics for the result."""
        if not self._result:
            return

        stats = {
            "total_desks": self._result.total_desks,
            "total_students": self._result.total_students,
            "cross_stage_pairs": self._result.cross_stage_count,
            "same_stage_pairs": self._result.same_stage_count,
            "single_student_desks": sum(
                1 for d in self._result.desks if d.student_count == 1
            ),
            "by_column": {col.value: 0 for col in Column},
            "by_stage": {}
        }

        for desk in self._result.desks:
            stats["by_column"][desk.column.value] += 1

            if desk.student_a:
                stage = desk.student_a.stage
                stats["by_stage"][stage] = stats["by_stage"].get(stage, 0) + 1

            if desk.student_b:
                stage = desk.student_b.stage
                stats["by_stage"][stage] = stats["by_stage"].get(stage, 0) + 1

        self._result.stats = stats

    def get_desks_by_column(self) -> Dict[str, List[Desk]]:
        """Get desks organized by column."""
        if not self._result:
            return {}

        by_column = {col.value: [] for col in Column}
        for desk in self._result.desks:
            by_column[desk.column.value].append(desk)
        return by_column

    def get_desks_by_row(self) -> Dict[int, List[Desk]]:
        """Get desks organized by row."""
        if not self._result:
            return {}

        by_row: Dict[int, List[Desk]] = {}
        for desk in self._result.desks:
            if desk.row not in by_row:
                by_row[desk.row] = []
            by_row[desk.row].append(desk)

        # Sort each row by column order
        column_order = {Column.RIGHT: 0, Column.MIDDLE: 1, Column.LEFT: 2}
        for row in by_row:
            by_row[row].sort(key=lambda d: column_order[d.column])

        return by_row

    def get_seating_grid(self) -> List[List[Optional[Desk]]]:
        """
        Get seating as a 2D grid (rows x columns).

        Returns:
            List of rows, each containing 3 desks (or None for empty slots)
        """
        if not self._result or not self._result.desks:
            return []

        by_row = self.get_desks_by_row()
        max_row = max(by_row.keys()) if by_row else 0

        grid = []
        for row_num in range(1, max_row + 1):
            row_desks = by_row.get(row_num, [])

            # Create row with 3 slots
            row = [None, None, None]  # Right, Middle, Left
            column_index = {Column.RIGHT: 0, Column.MIDDLE: 1, Column.LEFT: 2}

            for desk in row_desks:
                idx = column_index[desk.column]
                row[idx] = desk

            grid.append(row)

        return grid

    def clear(self) -> None:
        """Clear the current result."""
        self._result = None


# Convenience function for simple usage
def generate_seating(student_pools: Dict[str, List[str]]) -> Tuple[bool, str, SeatingResult]:
    """
    Generate anti-cheating seating arrangement.

    Args:
        student_pools: Dict mapping stage names to student name lists

    Returns:
        Tuple of (success, message, result)
    """
    algorithm = ZigzagSeatingAlgorithm()
    return algorithm.generate(student_pools)
