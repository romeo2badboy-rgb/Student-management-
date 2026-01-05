"""
Seating Algorithm Module
Implements smart seating arrangement with cross-stage pairing.
"""

import random
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass


@dataclass
class DeskAssignment:
    """Represents a single desk assignment."""
    desk_number: int
    column: str  # "Left", "Middle", "Right"
    student_a: str
    student_a_stage: str
    student_b: Optional[str]
    student_b_stage: Optional[str]


class SeatingAlgorithm:
    """
    Implements the smart seating algorithm that:
    - Pairs students from different stages at each desk
    - Randomizes students within each stage
    - Organizes desks into 3 columns (Left, Middle, Right)
    """

    COLUMNS = ["Left", "Middle", "Right"]

    def __init__(self):
        """Initialize the seating algorithm."""
        self._assignments: List[DeskAssignment] = []
        self._unassigned: List[Tuple[str, str]] = []  # (name, stage)

    @property
    def assignments(self) -> List[DeskAssignment]:
        """Get the list of desk assignments."""
        return self._assignments.copy()

    @property
    def unassigned_students(self) -> List[Tuple[str, str]]:
        """Get list of unassigned students (name, stage)."""
        return self._unassigned.copy()

    def generate_seating(self, student_pools: Dict[str, List[str]]) -> Tuple[bool, str]:
        """
        Generate seating arrangements from student pools.

        Args:
            student_pools: Dictionary with stage names as keys and student lists as values

        Returns:
            Tuple of (success: bool, message: str)
        """
        self._assignments = []
        self._unassigned = []

        # Filter out empty stages and create shuffled copies
        active_pools: Dict[str, List[str]] = {}
        for stage, students in student_pools.items():
            if students:
                shuffled = students.copy()
                random.shuffle(shuffled)
                active_pools[stage] = shuffled

        if len(active_pools) < 2:
            return False, "At least 2 stages with students are required for mixed seating."

        # Create a flat list of (student_name, stage) tuples
        all_students: List[Tuple[str, str]] = []
        for stage, students in active_pools.items():
            for student in students:
                all_students.append((student, stage))

        if len(all_students) < 2:
            return False, "At least 2 students are required."

        # Generate pairings using the smart algorithm
        pairs = self._create_cross_stage_pairs(active_pools)

        # Assign pairs to desks with column distribution
        desk_number = 1
        column_index = 0

        for student_a, stage_a, student_b, stage_b in pairs:
            column = self.COLUMNS[column_index % 3]

            assignment = DeskAssignment(
                desk_number=desk_number,
                column=column,
                student_a=student_a,
                student_a_stage=stage_a,
                student_b=student_b,
                student_b_stage=stage_b
            )
            self._assignments.append(assignment)

            desk_number += 1
            column_index += 1

        # Check for any unassigned students
        total_assigned = len(pairs) * 2 - pairs.count(None)
        total_students = sum(len(s) for s in active_pools.values())

        if self._unassigned:
            return True, f"Seating generated with {len(self._assignments)} desks. {len(self._unassigned)} student(s) could not be paired with different stages."

        return True, f"Successfully generated seating for {total_students} students in {len(self._assignments)} desks."

    def _create_cross_stage_pairs(self, pools: Dict[str, List[str]]) -> List[Tuple[str, str, Optional[str], Optional[str]]]:
        """
        Create pairs of students from different stages.

        Args:
            pools: Dictionary of stage -> shuffled student list

        Returns:
            List of tuples: (student_a, stage_a, student_b, stage_b)
        """
        pairs = []

        # Create working copies with stage tracking
        stage_queues: Dict[str, List[str]] = {stage: students.copy() for stage, students in pools.items()}
        stages = list(stage_queues.keys())

        # Sort stages by number of students (descending) for better distribution
        stages.sort(key=lambda s: len(stage_queues[s]), reverse=True)

        while True:
            # Remove empty stages
            stages = [s for s in stages if stage_queues[s]]

            if len(stages) == 0:
                break

            if len(stages) == 1:
                # Only one stage left - these students cannot be paired with different stages
                remaining_stage = stages[0]
                remaining_students = stage_queues[remaining_stage]

                # Pair remaining students from same stage (mark as same-stage pairs)
                while len(remaining_students) >= 2:
                    student_a = remaining_students.pop(0)
                    student_b = remaining_students.pop(0)
                    pairs.append((student_a, remaining_stage, student_b, remaining_stage))
                    self._unassigned.append((student_a, remaining_stage))
                    self._unassigned.append((student_b, remaining_stage))

                # Handle odd student
                if remaining_students:
                    student = remaining_students.pop(0)
                    pairs.append((student, remaining_stage, None, None))
                    self._unassigned.append((student, remaining_stage))
                break

            # Pick two different stages for pairing
            # Use round-robin style selection for fairness
            stage_a = stages[0]
            stage_b = stages[1]

            student_a = stage_queues[stage_a].pop(0)
            student_b = stage_queues[stage_b].pop(0)

            pairs.append((student_a, stage_a, student_b, stage_b))

            # Rotate stages for next iteration (for balanced distribution)
            stages = stages[1:] + [stages[0]]
            # Re-sort to handle varying lengths
            stages.sort(key=lambda s: len(stage_queues[s]) if stage_queues[s] else 0, reverse=True)

        return pairs

    def get_statistics(self) -> Dict:
        """
        Get statistics about the current seating arrangement.

        Returns:
            Dictionary with various statistics
        """
        if not self._assignments:
            return {"total_desks": 0, "total_students": 0}

        stats = {
            "total_desks": len(self._assignments),
            "total_students": 0,
            "by_column": {"Left": 0, "Middle": 0, "Right": 0},
            "cross_stage_pairs": 0,
            "same_stage_pairs": 0,
            "single_student_desks": 0
        }

        for assignment in self._assignments:
            stats["by_column"][assignment.column] += 1
            stats["total_students"] += 1

            if assignment.student_b:
                stats["total_students"] += 1
                if assignment.student_a_stage != assignment.student_b_stage:
                    stats["cross_stage_pairs"] += 1
                else:
                    stats["same_stage_pairs"] += 1
            else:
                stats["single_student_desks"] += 1

        return stats

    def get_assignments_by_column(self) -> Dict[str, List[DeskAssignment]]:
        """
        Get assignments organized by column.

        Returns:
            Dictionary with column names as keys and assignment lists as values
        """
        by_column = {"Left": [], "Middle": [], "Right": []}
        for assignment in self._assignments:
            by_column[assignment.column].append(assignment)
        return by_column

    def clear(self) -> None:
        """Clear all assignments."""
        self._assignments = []
        self._unassigned = []
