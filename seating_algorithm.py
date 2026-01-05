"""
Advanced Seating Algorithm Module
Implements the Anti-Cheating Zigzag Seating Pattern with strict stage alternation.
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
    is_empty: bool = False  # Flag for empty seats

    def __str__(self) -> str:
        if self.is_empty:
            return "--- Empty ---"
        return f"{self.name} ({self.stage})"

    @classmethod
    def empty_seat(cls) -> 'Student':
        """Create an empty seat placeholder."""
        return cls(name="--- Empty ---", stage="N/A", is_empty=True)


@dataclass
class Desk:
    """Represents a desk with two seats."""
    number: int
    column: Column
    row: int
    seat_a_number: int = 0  # Absolute seat number
    seat_b_number: int = 0
    student_a: Optional[Student] = None
    student_b: Optional[Student] = None

    @property
    def is_full(self) -> bool:
        """Check if desk has both students (including empty placeholders)."""
        return self.student_a is not None and self.student_b is not None

    @property
    def is_empty(self) -> bool:
        """Check if desk is completely empty."""
        return self.student_a is None and self.student_b is None

    @property
    def has_real_students(self) -> bool:
        """Check if desk has at least one real student."""
        has_a = self.student_a is not None and not self.student_a.is_empty
        has_b = self.student_b is not None and not self.student_b.is_empty
        return has_a or has_b

    @property
    def is_cross_stage(self) -> bool:
        """Check if desk has students from different stages."""
        if not self.is_full:
            return False
        if self.student_a.is_empty or self.student_b.is_empty:
            return False
        return self.student_a.stage != self.student_b.stage

    @property
    def student_count(self) -> int:
        """Get number of real students at this desk."""
        count = 0
        if self.student_a and not self.student_a.is_empty:
            count += 1
        if self.student_b and not self.student_b.is_empty:
            count += 1
        return count


@dataclass
class SeatingResult:
    """Contains the complete seating arrangement result."""
    desks: List[Desk] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    stats: Dict = field(default_factory=dict)
    stage_order: List[str] = field(default_factory=list)  # Track stage pattern used

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
        return sum(1 for desk in self.desks if desk.is_full and
                   desk.has_real_students and not desk.is_cross_stage and
                   not (desk.student_a.is_empty or desk.student_b.is_empty))

    @property
    def empty_seats_count(self) -> int:
        count = 0
        for desk in self.desks:
            if desk.student_a and desk.student_a.is_empty:
                count += 1
            if desk.student_b and desk.student_b.is_empty:
                count += 1
        return count


class ZigzagSeatingAlgorithm:
    """
    Advanced Anti-Cheating Seating Algorithm.

    Implements STRICT alternating pattern:
    - 2 stages: (1, 2, 1, 2, 1, 2, ...)
    - 3 stages: (1, 2, 3, 1, 2, 3, ...)

    Features:
    - Students from different stages sit at the same desk
    - Students are randomized within their stages
    - Desks are arranged in 3 columns (Right, Middle, Left)
    - Empty seats handled gracefully with placeholders
    """

    COLUMNS = [Column.RIGHT, Column.MIDDLE, Column.LEFT]
    DESKS_PER_ROW = 3

    def __init__(self):
        """Initialize the algorithm."""
        self._result: Optional[SeatingResult] = None

    @property
    def result(self) -> Optional[SeatingResult]:
        """Get the last seating result."""
        return self._result

    def generate(self, student_pools: Dict[str, List[str]]) -> Tuple[bool, str, SeatingResult]:
        """
        Generate anti-cheating seating arrangement with strict alternation.

        Args:
            student_pools: Dict mapping stage names to student name lists

        Returns:
            Tuple of (success, message, result)
        """
        self._result = SeatingResult()

        try:
            # Filter and validate pools
            valid_pools = {k: v for k, v in student_pools.items() if v}

            if len(valid_pools) == 0:
                return False, "No students found in any stage.", self._result

            if len(valid_pools) < 2:
                # Allow single stage but warn
                self._result.warnings.append(
                    "Only one stage has students. Cross-stage pairing not possible."
                )

            # Create ordered list of stages (sorted for consistency)
            stage_order = sorted(valid_pools.keys())
            self._result.stage_order = stage_order

            # Create Student objects and shuffle within each stage
            stage_queues: Dict[str, List[Student]] = {}
            total_students = 0

            for stage in stage_order:
                names = valid_pools[stage]
                students = [Student(name=name, stage=stage) for name in names]
                random.shuffle(students)  # Randomize within stage
                stage_queues[stage] = students
                total_students += len(students)

            if total_students == 0:
                return False, "No students to arrange.", self._result

            # Generate seating with strict alternation
            desks = self._create_strict_alternating_seating(stage_queues, stage_order)
            self._result.desks = desks
            self._calculate_stats()

            # Generate summary message
            message = (
                f"Generated {self._result.total_desks} desks for "
                f"{self._result.total_students} students. "
                f"Pattern: {' → '.join(stage_order)}"
            )

            if self._result.empty_seats_count > 0:
                self._result.warnings.append(
                    f"{self._result.empty_seats_count} empty seat(s) added due to odd student count."
                )

            return True, message, self._result

        except Exception as e:
            return False, f"Algorithm error: {str(e)}", self._result

    def _create_strict_alternating_seating(
        self,
        stage_queues: Dict[str, List[Student]],
        stage_order: List[str]
    ) -> List[Desk]:
        """
        Create seating with STRICT stage alternation pattern.

        Pattern examples:
        - 2 stages: Seat 1=Stage1, Seat 2=Stage2, Seat 3=Stage1, Seat 4=Stage2...
        - 3 stages: Seat 1=Stage1, Seat 2=Stage2, Seat 3=Stage3, Seat 4=Stage1...
        """
        desks: List[Desk] = []

        # Create a circular iterator for stages
        num_stages = len(stage_order)

        # Collect ALL students in strict alternating order
        all_students_ordered: List[Student] = []
        stage_indices = {stage: 0 for stage in stage_order}  # Track position in each queue

        # Calculate total students
        total_students = sum(len(q) for q in stage_queues.values())

        # Build the alternating sequence
        current_stage_idx = 0
        students_placed = 0
        max_iterations = total_students * num_stages * 2  # Safety limit
        iterations = 0

        while students_placed < total_students and iterations < max_iterations:
            iterations += 1
            stage = stage_order[current_stage_idx % num_stages]
            queue = stage_queues[stage]
            idx = stage_indices[stage]

            if idx < len(queue):
                # Add student from this stage
                all_students_ordered.append(queue[idx])
                stage_indices[stage] += 1
                students_placed += 1

            # Move to next stage in rotation
            current_stage_idx += 1

            # Check if we've exhausted all stages in this rotation
            if current_stage_idx % num_stages == 0:
                # Check if all stages are exhausted
                all_exhausted = all(
                    stage_indices[s] >= len(stage_queues[s])
                    for s in stage_order
                )
                if all_exhausted:
                    break

        # Handle odd number - add empty seat placeholder
        if len(all_students_ordered) % 2 == 1:
            all_students_ordered.append(Student.empty_seat())

        # Now pair students into desks (2 per desk)
        desk_number = 1
        seat_number = 1
        row = 1
        col_index = 0

        for i in range(0, len(all_students_ordered), 2):
            student_a = all_students_ordered[i] if i < len(all_students_ordered) else None
            student_b = all_students_ordered[i + 1] if i + 1 < len(all_students_ordered) else None

            # Safety check
            if student_a is None:
                break

            column = self.COLUMNS[col_index % 3]

            desk = Desk(
                number=desk_number,
                column=column,
                row=row,
                seat_a_number=seat_number,
                seat_b_number=seat_number + 1,
                student_a=student_a,
                student_b=student_b if student_b else Student.empty_seat()
            )
            desks.append(desk)

            desk_number += 1
            seat_number += 2
            col_index += 1

            # Move to next row after 3 columns
            if col_index % 3 == 0:
                row += 1

        return desks

    def _calculate_stats(self) -> None:
        """Calculate statistics for the result."""
        if not self._result:
            return

        stats = {
            "total_desks": self._result.total_desks,
            "total_students": self._result.total_students,
            "cross_stage_pairs": self._result.cross_stage_count,
            "same_stage_pairs": self._result.same_stage_count,
            "empty_seats": self._result.empty_seats_count,
            "single_student_desks": sum(
                1 for d in self._result.desks
                if d.student_count == 1
            ),
            "by_column": {col.value: 0 for col in Column},
            "by_stage": {},
            "stage_pattern": self._result.stage_order
        }

        for desk in self._result.desks:
            stats["by_column"][desk.column.value] += 1

            if desk.student_a and not desk.student_a.is_empty:
                stage = desk.student_a.stage
                stats["by_stage"][stage] = stats["by_stage"].get(stage, 0) + 1

            if desk.student_b and not desk.student_b.is_empty:
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
            by_row[row].sort(key=lambda d: column_order.get(d.column, 0))

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
        if not by_row:
            return []

        max_row = max(by_row.keys())

        grid = []
        for row_num in range(1, max_row + 1):
            row_desks = by_row.get(row_num, [])

            # Create row with 3 slots
            row = [None, None, None]  # Right, Middle, Left
            column_index = {Column.RIGHT: 0, Column.MIDDLE: 1, Column.LEFT: 2}

            for desk in row_desks:
                idx = column_index.get(desk.column, 0)
                if 0 <= idx < 3:
                    row[idx] = desk

            grid.append(row)

        return grid

    def clear(self) -> None:
        """Clear the current result."""
        self._result = None


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
