"""
Advanced Seating Algorithm Module v3.0
خوارزمية توزيع المقاعد المتقدمة

Features:
- الترتيب الشامل: خلط جميع الأقسام والمراحل
- تقسيم على قاعات: كل قاعة 18 مقعد (36 طالب)
- منع الغش: عدم جلوس طالبين من نفس المرحلة/القسم معاً
- دعم القاعات المتعددة
"""

import random
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
from enum import Enum
import math


class Column(Enum):
    """Exam hall column positions."""
    RIGHT = "يمين"
    MIDDLE = "وسط"
    LEFT = "يسار"


@dataclass
class Student:
    """Represents a student with their stage and department information."""
    name: str
    stage: str
    department: str = ""  # القسم
    is_empty: bool = False

    def __str__(self) -> str:
        if self.is_empty:
            return "--- فارغ ---"
        if self.department:
            return f"{self.name} ({self.stage[:3]}-{self.department[:3]})"
        return f"{self.name} ({self.stage})"

    @classmethod
    def empty_seat(cls) -> 'Student':
        """Create an empty seat placeholder."""
        return cls(name="--- فارغ ---", stage="N/A", department="", is_empty=True)

    @property
    def group_key(self) -> str:
        """مفتاح المجموعة للمقارنة (المرحلة + القسم)"""
        return f"{self.stage}|{self.department}"


@dataclass
class Desk:
    """Represents a desk with two seats."""
    number: int
    column: Column
    row: int
    seat_a_number: int = 0
    seat_b_number: int = 0
    student_a: Optional[Student] = None
    student_b: Optional[Student] = None

    @property
    def is_full(self) -> bool:
        return self.student_a is not None and self.student_b is not None

    @property
    def is_empty(self) -> bool:
        return self.student_a is None and self.student_b is None

    @property
    def has_real_students(self) -> bool:
        has_a = self.student_a is not None and not self.student_a.is_empty
        has_b = self.student_b is not None and not self.student_b.is_empty
        return has_a or has_b

    @property
    def is_cross_stage(self) -> bool:
        """هل الطالبان من مراحل/أقسام مختلفة؟"""
        if not self.is_full:
            return False
        if self.student_a.is_empty or self.student_b.is_empty:
            return False
        # مختلفين إذا كانوا من مرحلة مختلفة أو قسم مختلف
        return self.student_a.group_key != self.student_b.group_key

    @property
    def student_count(self) -> int:
        count = 0
        if self.student_a and not self.student_a.is_empty:
            count += 1
        if self.student_b and not self.student_b.is_empty:
            count += 1
        return count


@dataclass
class Room:
    """قاعة امتحانية تحتوي على عدة مقاعد"""
    number: int  # رقم القاعة
    desks: List[Desk] = field(default_factory=list)
    max_desks: int = 18  # عدد المقاعد الأقصى

    @property
    def total_students(self) -> int:
        return sum(desk.student_count for desk in self.desks)

    @property
    def cross_stage_pairs(self) -> int:
        return sum(1 for desk in self.desks if desk.is_cross_stage)

    @property
    def is_full(self) -> bool:
        return len(self.desks) >= self.max_desks


@dataclass
class SeatingResult:
    """Contains the complete seating arrangement result."""
    desks: List[Desk] = field(default_factory=list)
    rooms: List[Room] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    stats: Dict = field(default_factory=dict)
    stage_order: List[str] = field(default_factory=list)
    is_comprehensive: bool = False  # هل هذا ترتيب شامل؟

    @property
    def total_desks(self) -> int:
        if self.rooms:
            return sum(len(room.desks) for room in self.rooms)
        return len(self.desks)

    @property
    def total_students(self) -> int:
        if self.rooms:
            return sum(room.total_students for room in self.rooms)
        return sum(desk.student_count for desk in self.desks)

    @property
    def cross_stage_count(self) -> int:
        if self.rooms:
            return sum(room.cross_stage_pairs for room in self.rooms)
        return sum(1 for desk in self.desks if desk.is_cross_stage)

    @property
    def same_stage_count(self) -> int:
        all_desks = self.desks if not self.rooms else [d for r in self.rooms for d in r.desks]
        return sum(1 for desk in all_desks if desk.is_full and
                   desk.has_real_students and not desk.is_cross_stage and
                   not (desk.student_a.is_empty or desk.student_b.is_empty))

    @property
    def empty_seats_count(self) -> int:
        all_desks = self.desks if not self.rooms else [d for r in self.rooms for d in r.desks]
        count = 0
        for desk in all_desks:
            if desk.student_a and desk.student_a.is_empty:
                count += 1
            if desk.student_b and desk.student_b.is_empty:
                count += 1
        return count

    @property
    def total_rooms(self) -> int:
        return len(self.rooms) if self.rooms else 1


class ZigzagSeatingAlgorithm:
    """
    خوارزمية التوزيع المتقدمة لمنع الغش

    Modes:
    1. توزيع قسم واحد (العادي)
    2. الترتيب الشامل: خلط جميع الأقسام والمراحل
    """

    COLUMNS = [Column.RIGHT, Column.MIDDLE, Column.LEFT]
    DESKS_PER_ROW = 3
    DESKS_PER_ROOM = 18  # كل قاعة 18 مقعد = 36 طالب

    def __init__(self):
        self._result: Optional[SeatingResult] = None

    @property
    def result(self) -> Optional[SeatingResult]:
        return self._result

    def generate(self, student_pools: Dict[str, List[str]]) -> Tuple[bool, str, SeatingResult]:
        """
        توليد توزيع المقاعد لقسم واحد (الطريقة القديمة)
        """
        self._result = SeatingResult()

        try:
            valid_pools = {k: v for k, v in student_pools.items() if v}

            if len(valid_pools) == 0:
                return False, "لا يوجد طلاب في أي مرحلة", self._result

            if len(valid_pools) < 2:
                self._result.warnings.append(
                    "مرحلة واحدة فقط بها طلاب. لا يمكن منع الغش بشكل كامل"
                )

            stage_order = sorted(valid_pools.keys())
            self._result.stage_order = stage_order

            stage_queues: Dict[str, List[Student]] = {}
            total_students = 0

            for stage in stage_order:
                names = valid_pools[stage]
                students = [Student(name=name, stage=stage) for name in names]
                random.shuffle(students)
                stage_queues[stage] = students
                total_students += len(students)

            if total_students == 0:
                return False, "لا يوجد طلاب للتوزيع", self._result

            desks = self._create_strict_alternating_seating(stage_queues, stage_order)
            self._result.desks = desks
            self._calculate_stats()

            message = (
                f"تم توزيع {self._result.total_students} طالب على "
                f"{self._result.total_desks} مقعد. "
                f"النمط: {' ← '.join(stage_order)}"
            )

            if self._result.empty_seats_count > 0:
                self._result.warnings.append(
                    f"تم إضافة {self._result.empty_seats_count} مقعد فارغ بسبب العدد الفردي"
                )

            return True, message, self._result

        except Exception as e:
            return False, f"خطأ في الخوارزمية: {str(e)}", self._result

    def generate_comprehensive(
        self,
        all_students: List[Tuple[str, str, str]],
        desks_per_room: int = 18
    ) -> Tuple[bool, str, SeatingResult]:
        """
        الترتيب الشامل - خلط جميع الأقسام والمراحل

        Args:
            all_students: قائمة [(اسم_الطالب, المرحلة, القسم), ...]
            desks_per_room: عدد المقاعد في كل قاعة (الافتراضي 18)

        Returns:
            Tuple of (success, message, result)
        """
        self._result = SeatingResult()
        self._result.is_comprehensive = True
        self.DESKS_PER_ROOM = desks_per_room

        try:
            if not all_students:
                return False, "لا يوجد طلاب للتوزيع", self._result

            # إنشاء كائنات الطلاب
            students: List[Student] = []
            groups: Dict[str, List[Student]] = {}  # تجميع حسب (المرحلة + القسم)

            for name, stage, dept in all_students:
                student = Student(name=name, stage=stage, department=dept)
                students.append(student)

                group_key = student.group_key
                if group_key not in groups:
                    groups[group_key] = []
                groups[group_key].append(student)

            total_students = len(students)
            num_groups = len(groups)

            if num_groups < 2:
                self._result.warnings.append(
                    "جميع الطلاب من نفس المرحلة والقسم! لا يمكن منع الغش"
                )

            # خلط الطلاب داخل كل مجموعة
            for group_students in groups.values():
                random.shuffle(group_students)

            # إنشاء قائمة الطلاب المُرتبة بالتناوب
            ordered_students = self._create_alternating_list(groups)

            # إضافة مقعد فارغ إذا كان العدد فردياً
            if len(ordered_students) % 2 == 1:
                ordered_students.append(Student.empty_seat())

            # حساب عدد القاعات المطلوبة
            total_pairs = len(ordered_students) // 2
            students_per_room = desks_per_room * 2  # طالبين لكل مقعد
            num_rooms = math.ceil(total_students / students_per_room)

            # توزيع على القاعات
            rooms: List[Room] = []
            student_idx = 0
            seat_number = 1

            for room_num in range(1, num_rooms + 1):
                room = Room(number=room_num, max_desks=desks_per_room)
                desk_number = 1
                row = 1
                col_idx = 0

                while len(room.desks) < desks_per_room and student_idx < len(ordered_students) - 1:
                    student_a = ordered_students[student_idx]
                    student_b = ordered_students[student_idx + 1] if student_idx + 1 < len(ordered_students) else Student.empty_seat()

                    column = self.COLUMNS[col_idx % 3]

                    desk = Desk(
                        number=desk_number,
                        column=column,
                        row=row,
                        seat_a_number=seat_number,
                        seat_b_number=seat_number + 1,
                        student_a=student_a,
                        student_b=student_b
                    )

                    room.desks.append(desk)
                    student_idx += 2
                    seat_number += 2
                    desk_number += 1
                    col_idx += 1

                    if col_idx % 3 == 0:
                        row += 1

                rooms.append(room)

            self._result.rooms = rooms
            self._result.stage_order = list(groups.keys())
            self._calculate_stats_comprehensive()

            message = (
                f"تم التوزيع الشامل: {self._result.total_students} طالب "
                f"على {len(rooms)} قاعة "
                f"({self._result.total_desks} مقعد)"
            )

            if self._result.empty_seats_count > 0:
                self._result.warnings.append(
                    f"تم إضافة {self._result.empty_seats_count} مقعد فارغ"
                )

            # إحصائية منع الغش
            cross_percentage = (self._result.cross_stage_count / max(1, self._result.total_desks)) * 100
            if cross_percentage < 80:
                self._result.warnings.append(
                    f"نسبة منع الغش: {cross_percentage:.1f}% (المثالي > 80%)"
                )

            return True, message, self._result

        except Exception as e:
            return False, f"خطأ في الترتيب الشامل: {str(e)}", self._result

    def _create_alternating_list(self, groups: Dict[str, List[Student]]) -> List[Student]:
        """
        إنشاء قائمة طلاب بالتناوب بين المجموعات
        بحيث لا يجلس طالبان من نفس المجموعة معاً
        """
        result: List[Student] = []
        group_keys = list(groups.keys())
        num_groups = len(group_keys)

        if num_groups == 0:
            return result

        if num_groups == 1:
            # مجموعة واحدة فقط - لا يمكن التناوب
            return groups[group_keys[0]].copy()

        # إنشاء فهرس لكل مجموعة
        indices = {key: 0 for key in group_keys}
        total_students = sum(len(g) for g in groups.values())

        current_group_idx = 0
        last_group_key = None

        iterations = 0
        max_iterations = total_students * num_groups * 2

        while len(result) < total_students and iterations < max_iterations:
            iterations += 1

            # البحث عن مجموعة مختلفة عن السابقة
            found = False
            attempts = 0

            while attempts < num_groups:
                group_key = group_keys[(current_group_idx + attempts) % num_groups]
                idx = indices[group_key]

                if idx < len(groups[group_key]) and group_key != last_group_key:
                    # وجدنا طالباً من مجموعة مختلفة
                    result.append(groups[group_key][idx])
                    indices[group_key] += 1
                    last_group_key = group_key
                    current_group_idx = (current_group_idx + attempts + 1) % num_groups
                    found = True
                    break

                attempts += 1

            if not found:
                # لم نجد مجموعة مختلفة - نأخذ أي طالب متبقي
                for group_key in group_keys:
                    idx = indices[group_key]
                    if idx < len(groups[group_key]):
                        result.append(groups[group_key][idx])
                        indices[group_key] += 1
                        last_group_key = group_key
                        break

        return result

    def _create_strict_alternating_seating(
        self,
        stage_queues: Dict[str, List[Student]],
        stage_order: List[str]
    ) -> List[Desk]:
        """إنشاء توزيع بالتناوب الصارم بين المراحل"""
        desks: List[Desk] = []

        num_stages = len(stage_order)
        all_students_ordered: List[Student] = []
        stage_indices = {stage: 0 for stage in stage_order}

        total_students = sum(len(q) for q in stage_queues.values())

        current_stage_idx = 0
        students_placed = 0
        max_iterations = total_students * num_stages * 2
        iterations = 0

        while students_placed < total_students and iterations < max_iterations:
            iterations += 1
            stage = stage_order[current_stage_idx % num_stages]
            queue = stage_queues[stage]
            idx = stage_indices[stage]

            if idx < len(queue):
                all_students_ordered.append(queue[idx])
                stage_indices[stage] += 1
                students_placed += 1

            current_stage_idx += 1

            if current_stage_idx % num_stages == 0:
                all_exhausted = all(
                    stage_indices[s] >= len(stage_queues[s])
                    for s in stage_order
                )
                if all_exhausted:
                    break

        if len(all_students_ordered) % 2 == 1:
            all_students_ordered.append(Student.empty_seat())

        desk_number = 1
        seat_number = 1
        row = 1
        col_index = 0

        for i in range(0, len(all_students_ordered), 2):
            student_a = all_students_ordered[i] if i < len(all_students_ordered) else None
            student_b = all_students_ordered[i + 1] if i + 1 < len(all_students_ordered) else None

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

            if col_index % 3 == 0:
                row += 1

        return desks

    def _calculate_stats(self) -> None:
        """حساب الإحصائيات للتوزيع العادي"""
        if not self._result:
            return

        stats = {
            "total_desks": self._result.total_desks,
            "total_students": self._result.total_students,
            "cross_stage_pairs": self._result.cross_stage_count,
            "same_stage_pairs": self._result.same_stage_count,
            "empty_seats": self._result.empty_seats_count,
            "total_rooms": 1,
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

    def _calculate_stats_comprehensive(self) -> None:
        """حساب الإحصائيات للترتيب الشامل"""
        if not self._result:
            return

        stats = {
            "total_desks": self._result.total_desks,
            "total_students": self._result.total_students,
            "cross_stage_pairs": self._result.cross_stage_count,
            "same_stage_pairs": self._result.same_stage_count,
            "empty_seats": self._result.empty_seats_count,
            "total_rooms": len(self._result.rooms),
            "by_room": {},
            "by_stage": {},
            "by_department": {},
            "is_comprehensive": True
        }

        for room in self._result.rooms:
            room_stats = {
                "desks": len(room.desks),
                "students": room.total_students,
                "cross_stage": room.cross_stage_pairs
            }
            stats["by_room"][room.number] = room_stats

            for desk in room.desks:
                if desk.student_a and not desk.student_a.is_empty:
                    stage = desk.student_a.stage
                    dept = desk.student_a.department
                    stats["by_stage"][stage] = stats["by_stage"].get(stage, 0) + 1
                    if dept:
                        stats["by_department"][dept] = stats["by_department"].get(dept, 0) + 1

                if desk.student_b and not desk.student_b.is_empty:
                    stage = desk.student_b.stage
                    dept = desk.student_b.department
                    stats["by_stage"][stage] = stats["by_stage"].get(stage, 0) + 1
                    if dept:
                        stats["by_department"][dept] = stats["by_department"].get(dept, 0) + 1

        self._result.stats = stats

    def get_desks_by_column(self) -> Dict[str, List[Desk]]:
        """الحصول على المقاعد مرتبة بالأعمدة"""
        if not self._result:
            return {}

        all_desks = self._result.desks if not self._result.rooms else [
            d for r in self._result.rooms for d in r.desks
        ]

        by_column = {col.value: [] for col in Column}
        for desk in all_desks:
            by_column[desk.column.value].append(desk)
        return by_column

    def get_desks_by_row(self) -> Dict[int, List[Desk]]:
        """الحصول على المقاعد مرتبة بالصفوف"""
        if not self._result:
            return {}

        all_desks = self._result.desks if not self._result.rooms else [
            d for r in self._result.rooms for d in r.desks
        ]

        by_row: Dict[int, List[Desk]] = {}
        for desk in all_desks:
            if desk.row not in by_row:
                by_row[desk.row] = []
            by_row[desk.row].append(desk)

        column_order = {Column.RIGHT: 0, Column.MIDDLE: 1, Column.LEFT: 2}
        for row in by_row:
            by_row[row].sort(key=lambda d: column_order.get(d.column, 0))

        return by_row

    def get_seating_grid(self) -> List[List[Optional[Desk]]]:
        """الحصول على التوزيع كشبكة ثنائية الأبعاد"""
        if not self._result:
            return []

        all_desks = self._result.desks if not self._result.rooms else [
            d for r in self._result.rooms for d in r.desks
        ]

        if not all_desks:
            return []

        by_row = self.get_desks_by_row()
        if not by_row:
            return []

        max_row = max(by_row.keys())

        grid = []
        for row_num in range(1, max_row + 1):
            row_desks = by_row.get(row_num, [])
            row = [None, None, None]
            column_index = {Column.RIGHT: 0, Column.MIDDLE: 1, Column.LEFT: 2}

            for desk in row_desks:
                idx = column_index.get(desk.column, 0)
                if 0 <= idx < 3:
                    row[idx] = desk

            grid.append(row)

        return grid

    def clear(self) -> None:
        """مسح النتيجة الحالية"""
        self._result = None


def generate_seating(student_pools: Dict[str, List[str]]) -> Tuple[bool, str, SeatingResult]:
    """توليد توزيع المقاعد"""
    algorithm = ZigzagSeatingAlgorithm()
    return algorithm.generate(student_pools)


def generate_comprehensive_seating(
    all_students: List[Tuple[str, str, str]],
    desks_per_room: int = 18
) -> Tuple[bool, str, SeatingResult]:
    """
    توليد الترتيب الشامل

    Args:
        all_students: [(اسم, مرحلة, قسم), ...]
        desks_per_room: عدد المقاعد في كل قاعة

    Returns:
        Tuple of (success, message, result)
    """
    algorithm = ZigzagSeatingAlgorithm()
    return algorithm.generate_comprehensive(all_students, desks_per_room)
