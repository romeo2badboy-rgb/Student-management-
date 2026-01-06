#!/usr/bin/env python3
"""
نظام توزيع مقاعد الامتحانات - الإصدار 3.0
University Exam Seating System - Arabic Version v3.0

المميزات:
- واجهة عربية بالكامل
- دعم متعدد الأقسام
- استيراد الطلاب من ملفات Word و Excel
- خوارزمية توزيع لمنع الغش
- تصدير تقارير احترافية
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Optional, Callable, Dict, Any
import os
import traceback

from data_manager import DataManager
from seating_algorithm import ZigzagSeatingAlgorithm, SeatingResult
from word_exporter import WordExporter


# ==================== إعدادات المظهر ====================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# نظام الألوان
COLORS = {
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_light": "#3B82F6",
    "secondary": "#6B7280",
    "success": "#059669",
    "success_hover": "#047857",
    "warning": "#D97706",
    "warning_hover": "#B45309",
    "danger": "#DC2626",
    "danger_hover": "#B91C1C",
    "bg_dark": "#111827",
    "bg_card": "#1F2937",
    "bg_hover": "#374151",
    "text": "#F9FAFB",
    "text_secondary": "#9CA3AF",
    "border": "#374151",
    "accent": "#7C3AED"
}

# مقاسات الواجهة
SCALE = {
    "sidebar_width": 220,
    "card_corner": 15,
    "button_height": 45,
    "entry_height": 45,
    "font_title": 26,
    "font_subtitle": 13,
    "font_normal": 14,
    "font_small": 12,
    "padding_large": 25,
    "padding_medium": 15,
    "padding_small": 10
}

# النصوص العربية
ARABIC_TEXT = {
    # العناوين الرئيسية
    "app_title": "نظام توزيع المقاعد",
    "app_subtitle": "إدارة الامتحانات",
    "version": "الإصدار 3.0",

    # القائمة الجانبية
    "nav_dashboard": "الرئيسية",
    "nav_departments": "الأقسام",
    "nav_students": "الطلاب",
    "nav_generate": "توليد التوزيع",
    "nav_export": "التصدير",
    "clear_all": "مسح جميع البيانات",
    "theme": "المظهر",
    "dark_mode": "الوضع الداكن",

    # لوحة التحكم
    "dashboard_title": "لوحة التحكم",
    "dashboard_subtitle": "نظرة عامة على النظام",
    "university_name": "اسم الجامعة/المدرسة",
    "enter_uni_name": "أدخل اسم المؤسسة...",
    "departments_count": "الأقسام",
    "students_count": "الطلاب",
    "stages_count": "المراحل",
    "desks_count": "المقاعد",
    "quick_actions": "إجراءات سريعة",
    "new_department": "قسم جديد",
    "import_students": "استيراد طلاب",
    "generate_seating": "توليد التوزيع",
    "export_report": "تصدير التقرير",
    "quick_guide": "دليل الاستخدام",

    # صفحة الأقسام
    "departments_title": "إدارة الأقسام",
    "departments_subtitle": "إضافة وحذف الأقسام",
    "create_department": "إنشاء قسم جديد",
    "dept_placeholder": "أدخل اسم القسم...",
    "create_btn": "إنشاء",
    "existing_departments": "الأقسام الحالية",
    "no_departments": "لا توجد أقسام\n\nأنشئ قسمك الأول للبدء",
    "delete": "حذف",
    "total": "المجموع",
    "students": "طالب",

    # صفحة الطلاب
    "students_title": "إدارة الطلاب",
    "students_subtitle": "استيراد وإدارة قوائم الطلاب",
    "import_from_file": "استيراد من ملف",
    "department": "القسم",
    "select_dept": "اختر القسم",
    "stage": "المرحلة",
    "stage_1": "المرحلة الأولى",
    "stage_2": "المرحلة الثانية",
    "stage_3": "المرحلة الثالثة",
    "import_word": "استيراد Word",
    "import_excel": "استيراد Excel",
    "clear_stage": "مسح المرحلة",
    "add_manually": "إضافة يدوية",
    "student_name": "اسم الطالب...",
    "add_student": "إضافة",
    "student_lists": "قوائم الطلاب",
    "students_label": "الطلاب:",
    "no_students": "لا يوجد طلاب في هذه المرحلة",
    "select_dept_first": "يرجى اختيار القسم أولاً",

    # صفحة التوليد
    "generate_title": "توليد التوزيع",
    "generate_subtitle": "إنشاء توزيع المقاعد بنمط مضاد للغش",
    "algorithm_info": "خوارزمية منع الغش",
    "algorithm_desc": "• التبديل الصارم: مرحلتان (1,2,1,2...) أو ثلاث (1,2,3,1,2,3...)\n• الطلاب مُرتبين عشوائياً داخل كل مرحلة\n• المقاعد في 3 أعمدة (يمين، وسط، يسار)\n• المقاعد الفارغة مُعلّمة عند الحاجة",
    "select_department": "اختر القسم:",
    "generate_btn": "توليد التوزيع",
    "results_title": "نتيجة التوزيع",
    "initial_message": "اختر قسماً واضغط 'توليد التوزيع' للبدء...",
    "pattern": "نمط التوزيع",
    "desk": "مقعد",
    "column": "عمود",
    "row": "صف",
    "seat_a": "مقعد أ",
    "seat_b": "مقعد ب",
    "student_a": "الطالب أ",
    "student_b": "الطالب ب",
    "empty_seat": "--- فارغ ---",
    "cross_stage": "أزواج مختلطة",
    "empty_seats": "مقاعد فارغة",

    # صفحة التصدير
    "export_title": "تصدير التقرير",
    "export_subtitle": "إنشاء مستند Word احترافي",
    "export_options": "خيارات التصدير",
    "room_number": "رقم القاعة",
    "director_name": "مدير الاعدادية",
    "exam_title": "عنوان الامتحان",
    "exam_date": "التاريخ",
    "date_placeholder": "اتركه فارغاً لتاريخ اليوم",
    "include_map": "تضمين خريطة المقاعد",
    "export_btn": "تصدير إلى Word",
    "preview": "معاينة",
    "no_data": "لا توجد بيانات للتصدير",
    "generate_first": "يرجى توليد التوزيع أولاً",
    "ready_export": "البيانات جاهزة للتصدير!",

    # الرسائل
    "warning": "تحذير",
    "error": "خطأ",
    "success": "نجاح",
    "confirm": "تأكيد",
    "confirm_delete": "هل تريد حذف القسم '{}'؟\n\nسيتم حذف جميع بيانات الطلاب.",
    "confirm_clear": "هل تريد مسح جميع البيانات؟\n\nسيتم حذف:\n• جميع الأقسام\n• جميع قوائم الطلاب\n• التوزيع الحالي",
    "data_cleared": "تم مسح جميع البيانات\nالنظام جاهز لجلسة جديدة",
    "export_success": "تم التصدير بنجاح",
    "open_file": "هل تريد فتح الملف؟",
    "no_dept_selected": "يرجى اختيار القسم",
    "no_students_stage": "لا يوجد طلاب في {} للمسح",
    "clear_confirm": "هل تريد مسح {} طالب من {}؟",
    "cleared_students": "تم مسح {} طالب من {}",
    "import_success": "تم الاستيراد بنجاح",
    "validation_error": "خطأ في التحقق",
    "generation_error": "خطأ في التوليد",
}


# ==================== معالج الأخطاء ====================

def safe_execute(func: Callable, error_title: str = "خطأ") -> Any:
    """تنفيذ آمن للدوال مع معالجة الأخطاء"""
    try:
        return func()
    except Exception as e:
        error_msg = str(e)
        print(f"Error: {error_msg}")
        print(traceback.format_exc())
        messagebox.showerror(error_title, f"حدث خطأ:\n{error_msg}")
        return None


# ==================== القائمة الجانبية ====================

class Sidebar(ctk.CTkFrame):
    """القائمة الجانبية للتنقل"""

    def __init__(self, master, nav_callback: Callable, clear_callback: Callable, **kwargs):
        super().__init__(master, width=SCALE["sidebar_width"], corner_radius=0,
                        fg_color=COLORS["bg_card"], **kwargs)

        self.nav_callback = nav_callback
        self.clear_callback = clear_callback
        self.buttons: Dict[str, ctk.CTkButton] = {}
        self.current_page = "dashboard"

        self.grid_propagate(False)
        self.pack_propagate(False)

        self._create_header()
        self._create_navigation()
        self._create_footer()

    def _create_header(self):
        """إنشاء رأس القائمة"""
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=15, pady=(25, 10))

        # أيقونة بسيطة بدلاً من إيموجي
        self.icon_frame = ctk.CTkFrame(
            self.header_frame,
            width=60, height=60,
            corner_radius=30,
            fg_color=COLORS["primary"]
        )
        self.icon_frame.pack()
        self.icon_frame.pack_propagate(False)

        self.icon_label = ctk.CTkLabel(
            self.icon_frame,
            text="EM",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="white"
        )
        self.icon_label.place(relx=0.5, rely=0.5, anchor="center")

        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text=ARABIC_TEXT["app_title"],
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.title_label.pack(pady=(12, 0))

        self.subtitle_label = ctk.CTkLabel(
            self.header_frame,
            text=ARABIC_TEXT["app_subtitle"],
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle_label.pack()

        # خط فاصل
        self.divider = ctk.CTkFrame(self, height=1, fg_color=COLORS["border"])
        self.divider.pack(fill="x", padx=20, pady=20)

    def _create_navigation(self):
        """إنشاء أزرار التنقل"""
        nav_items = [
            ("dashboard", ARABIC_TEXT["nav_dashboard"]),
            ("departments", ARABIC_TEXT["nav_departments"]),
            ("students", ARABIC_TEXT["nav_students"]),
            ("generate", ARABIC_TEXT["nav_generate"]),
            ("export", ARABIC_TEXT["nav_export"]),
        ]

        for key, text in nav_items:
            btn = ctk.CTkButton(
                self,
                text=text,
                font=ctk.CTkFont(size=14),
                height=46,
                corner_radius=10,
                fg_color="transparent",
                text_color=COLORS["text"],
                hover_color=COLORS["bg_hover"],
                anchor="center",
                command=lambda k=key: self._navigate(k)
            )
            btn.pack(fill="x", padx=12, pady=3)
            self.buttons[key] = btn

        self._highlight("dashboard")

        # مساحة فارغة
        self.spacer = ctk.CTkFrame(self, fg_color="transparent")
        self.spacer.pack(fill="both", expand=True)

    def _create_footer(self):
        """إنشاء ذيل القائمة"""
        # خط فاصل
        self.clear_divider = ctk.CTkFrame(self, height=1, fg_color=COLORS["border"])
        self.clear_divider.pack(fill="x", padx=20, pady=(10, 10))

        # زر مسح البيانات
        self.clear_btn = ctk.CTkButton(
            self,
            text=ARABIC_TEXT["clear_all"],
            font=ctk.CTkFont(size=12),
            height=38,
            corner_radius=8,
            fg_color=COLORS["danger"],
            hover_color=COLORS["danger_hover"],
            command=self._on_clear_all
        )
        self.clear_btn.pack(fill="x", padx=12, pady=(0, 10))

        # تبديل المظهر
        self.theme_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.theme_frame.pack(fill="x", padx=15, pady=(5, 8))

        self.theme_switch = ctk.CTkSwitch(
            self.theme_frame,
            text=ARABIC_TEXT["dark_mode"],
            font=ctk.CTkFont(size=11),
            command=self._toggle_theme,
            onvalue="dark",
            offvalue="light",
            height=22
        )
        self.theme_switch.pack(anchor="center")
        self.theme_switch.select()

        # الإصدار
        self.version_label = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["version"],
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_secondary"]
        )
        self.version_label.pack(pady=(5, 15))

    def _navigate(self, key: str):
        """معالجة النقر على زر التنقل"""
        try:
            self.current_page = key
            self._highlight(key)
            self.nav_callback(key)
        except Exception as e:
            print(f"Navigation error: {e}")

    def _highlight(self, key: str):
        """تمييز الزر النشط"""
        for btn_key, btn in self.buttons.items():
            if btn_key == key:
                btn.configure(fg_color=COLORS["primary"])
            else:
                btn.configure(fg_color="transparent")

    def _toggle_theme(self):
        """تبديل المظهر"""
        try:
            mode = self.theme_switch.get()
            ctk.set_appearance_mode(mode)
        except Exception as e:
            print(f"Theme toggle error: {e}")

    def _on_clear_all(self):
        """معالجة زر مسح البيانات"""
        self.clear_callback()


# ==================== صفحة لوحة التحكم ====================

class DashboardPage(ctk.CTkFrame):
    """الصفحة الرئيسية"""

    def __init__(self, master, data_manager: DataManager, nav_callback: Callable = None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.nav_callback = nav_callback

        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_frame.pack(fill="both", expand=True)

        self._create_header()
        self._create_university_card()
        self._create_stats_section()
        self._create_quick_actions()
        self._create_guide_section()

    def _create_header(self):
        """إنشاء رأس الصفحة"""
        self.title = ctk.CTkLabel(
            self.scroll_frame,
            text=ARABIC_TEXT["dashboard_title"],
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="e", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self.scroll_frame,
            text=ARABIC_TEXT["dashboard_subtitle"],
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="e", pady=(0, SCALE["padding_large"]))

    def _create_university_card(self):
        """إنشاء بطاقة اسم الجامعة"""
        self.uni_card = ctk.CTkFrame(self.scroll_frame, corner_radius=SCALE["card_corner"])
        self.uni_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.uni_label = ctk.CTkLabel(
            self.uni_card,
            text=ARABIC_TEXT["university_name"],
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"]
        )
        self.uni_label.pack(anchor="e", padx=20, pady=(15, 5))

        self.uni_entry = ctk.CTkEntry(
            self.uni_card,
            placeholder_text=ARABIC_TEXT["enter_uni_name"],
            height=SCALE["entry_height"],
            font=ctk.CTkFont(size=SCALE["font_normal"]),
            justify="right"
        )
        self.uni_entry.pack(fill="x", padx=20, pady=(0, 15))
        self.uni_entry.insert(0, "اسم المؤسسة")
        self.uni_entry.bind("<KeyRelease>", self._on_uni_change)

    def _create_stats_section(self):
        """إنشاء قسم الإحصائيات"""
        self.stats_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.stats_frame.pack(fill="x", pady=10)
        self.stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.stat_cards: Dict[str, ctk.CTkFrame] = {}
        stats_config = [
            ("departments", ARABIC_TEXT["departments_count"], "0", COLORS["primary"]),
            ("students", ARABIC_TEXT["students_count"], "0", COLORS["success"]),
            ("stages", ARABIC_TEXT["stages_count"], "0", COLORS["warning"]),
            ("desks", ARABIC_TEXT["desks_count"], "0", COLORS["accent"])
        ]

        for i, (key, label, value, color) in enumerate(stats_config):
            card = self._create_stat_card(self.stats_frame, label, value, color)
            card.grid(row=0, column=3-i, padx=8, pady=5, sticky="nsew")
            self.stat_cards[key] = card

    def _create_stat_card(self, parent, label: str, value: str, accent_color: str) -> ctk.CTkFrame:
        """إنشاء بطاقة إحصائية"""
        card = ctk.CTkFrame(parent, corner_radius=SCALE["card_corner"], height=130)
        card.pack_propagate(False)

        # شريط ملون في الأعلى
        accent = ctk.CTkFrame(card, height=4, fg_color=accent_color, corner_radius=2)
        accent.pack(fill="x", padx=12, pady=(12, 0))

        value_label = ctk.CTkLabel(
            card,
            text=value,
            font=ctk.CTkFont(size=36, weight="bold")
        )
        value_label.pack(pady=(15, 5))

        text_label = ctk.CTkLabel(
            card,
            text=label,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"]
        )
        text_label.pack(pady=(0, 12))

        card.value_label = value_label
        return card

    def _create_quick_actions(self):
        """إنشاء أزرار الإجراءات السريعة"""
        self.actions_label = ctk.CTkLabel(
            self.scroll_frame,
            text=ARABIC_TEXT["quick_actions"],
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.actions_label.pack(anchor="e", pady=(SCALE["padding_large"], SCALE["padding_medium"]))

        self.actions_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.actions_frame.pack(fill="x")

        actions = [
            (ARABIC_TEXT["export_report"], COLORS["accent"], lambda: self._quick_nav("export")),
            (ARABIC_TEXT["generate_seating"], COLORS["warning"], lambda: self._quick_nav("generate")),
            (ARABIC_TEXT["import_students"], COLORS["primary"], lambda: self._quick_nav("students")),
            (ARABIC_TEXT["new_department"], COLORS["success"], lambda: self._quick_nav("departments")),
        ]

        for i, (text, color, command) in enumerate(actions):
            btn = ctk.CTkButton(
                self.actions_frame,
                text=text,
                font=ctk.CTkFont(size=13, weight="bold"),
                height=48,
                width=160,
                fg_color=color,
                hover_color=self._darken_color(color),
                command=command
            )
            btn.grid(row=0, column=i, padx=8, pady=5)

    def _create_guide_section(self):
        """إنشاء قسم الدليل"""
        self.guide_label = ctk.CTkLabel(
            self.scroll_frame,
            text=ARABIC_TEXT["quick_guide"],
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.guide_label.pack(anchor="e", pady=(SCALE["padding_large"], SCALE["padding_medium"]))

        self.guide_card = ctk.CTkFrame(self.scroll_frame, corner_radius=SCALE["card_corner"])
        self.guide_card.pack(fill="both", expand=True)

        self.guide_text = ctk.CTkTextbox(
            self.guide_card,
            font=ctk.CTkFont(size=13),
            fg_color="transparent",
            height=200
        )
        self.guide_text.pack(fill="both", expand=True, padx=15, pady=15)
        self._update_guide_text()
        self.guide_text.configure(state="disabled")

    def _darken_color(self, hex_color: str) -> str:
        """تعتيم لون للتأثير عند التمرير"""
        try:
            hex_color = hex_color.lstrip('#')
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            r = max(0, int(r * 0.8))
            g = max(0, int(g * 0.8))
            b = max(0, int(b * 0.8))
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return hex_color

    def _quick_nav(self, page: str):
        """التنقل السريع"""
        if self.nav_callback:
            self.nav_callback(page)

    def _update_guide_text(self):
        """تحديث نص الدليل"""
        self.guide_text.configure(state="normal")
        self.guide_text.delete("1.0", "end")

        guide_text = """مرحباً بك في نظام توزيع مقاعد الامتحانات!

كيفية الاستخدام:
═══════════════════════════════════════════════

الخطوة 1: إنشاء الأقسام
   - اذهب إلى صفحة "الأقسام"
   - أدخل اسم القسم واضغط "إنشاء"

الخطوة 2: استيراد الطلاب
   - اذهب إلى صفحة "الطلاب"
   - اختر القسم والمرحلة
   - استورد من ملف Word أو Excel
   - أو أضف الطلاب يدوياً

الخطوة 3: توليد التوزيع
   - اذهب إلى صفحة "توليد التوزيع"
   - اختر القسم واضغط "توليد التوزيع"
   - سيتم ترتيب الطلاب بنمط مضاد للغش

الخطوة 4: تصدير التقرير
   - اذهب إلى صفحة "التصدير"
   - أدخل البيانات وصدّر إلى Word

نمط منع الغش:
   - مرحلتان: 1 ← 2 ← 1 ← 2...
   - ثلاث مراحل: 1 ← 2 ← 3 ← 1 ← 2 ← 3..."""

        self.guide_text.insert("1.0", guide_text)
        self.guide_text.configure(state="disabled")

    def _on_uni_change(self, event):
        """تحديث اسم الجامعة"""
        try:
            self.data_manager.university_name = self.uni_entry.get()
        except Exception as e:
            print(f"Error updating university name: {e}")

    def refresh(self):
        """تحديث الإحصائيات"""
        try:
            current_uni = self.data_manager.university_name
            if self.uni_entry.get() != current_uni:
                self.uni_entry.delete(0, "end")
                self.uni_entry.insert(0, current_uni)

            summary = self.data_manager.get_global_summary()

            dept_count = summary.get("total_departments", 0)
            self.stat_cards["departments"].value_label.configure(text=str(dept_count))

            total_students = sum(
                d.get("total_students", 0)
                for d in summary.get("departments", [])
            )
            self.stat_cards["students"].value_label.configure(text=str(total_students))

            active_stages = sum(
                sum(1 for v in d.get("stages", {}).values() if v > 0)
                for d in summary.get("departments", [])
            )
            self.stat_cards["stages"].value_label.configure(text=str(active_stages))

            desks_needed = (total_students + 1) // 2
            self.stat_cards["desks"].value_label.configure(text=str(desks_needed))
        except Exception as e:
            print(f"Error refreshing dashboard: {e}")


# ==================== صفحة الأقسام ====================

class DepartmentsPage(ctk.CTkFrame):
    """صفحة إدارة الأقسام"""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.dept_widgets = []

        self._create_header()
        self._create_add_section()
        self._create_list_section()

    def _create_header(self):
        """إنشاء رأس الصفحة"""
        self.title = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["departments_title"],
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="e", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["departments_subtitle"],
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="e", pady=(0, SCALE["padding_large"]))

    def _create_add_section(self):
        """إنشاء قسم الإضافة"""
        self.add_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.add_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.add_label = ctk.CTkLabel(
            self.add_card,
            text=ARABIC_TEXT["create_department"],
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.add_label.pack(anchor="e", padx=20, pady=(15, 10))

        self.add_frame = ctk.CTkFrame(self.add_card, fg_color="transparent")
        self.add_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.add_btn = ctk.CTkButton(
            self.add_frame,
            text=ARABIC_TEXT["create_btn"],
            font=ctk.CTkFont(size=SCALE["font_normal"], weight="bold"),
            height=SCALE["button_height"],
            width=120,
            fg_color=COLORS["success"],
            hover_color=COLORS["success_hover"],
            command=self._create_department
        )
        self.add_btn.pack(side="left")

        self.dept_entry = ctk.CTkEntry(
            self.add_frame,
            placeholder_text=ARABIC_TEXT["dept_placeholder"],
            height=SCALE["entry_height"],
            font=ctk.CTkFont(size=SCALE["font_normal"]),
            justify="right"
        )
        self.dept_entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
        self.dept_entry.bind("<Return>", lambda e: self._create_department())

    def _create_list_section(self):
        """إنشاء قسم القائمة"""
        self.list_label = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["existing_departments"],
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.list_label.pack(anchor="e", pady=(10, SCALE["padding_medium"]))

        self.list_frame = ctk.CTkScrollableFrame(self, corner_radius=SCALE["card_corner"])
        self.list_frame.pack(fill="both", expand=True)

    def _create_department(self):
        """إنشاء قسم جديد"""
        try:
            name = self.dept_entry.get().strip()
            if not name:
                messagebox.showwarning(ARABIC_TEXT["warning"], "يرجى إدخال اسم القسم")
                return

            success, message = self.data_manager.create_department(name)
            if success:
                self.dept_entry.delete(0, "end")
                self.refresh()
                messagebox.showinfo(ARABIC_TEXT["success"], f"تم إنشاء القسم: {name}")
            else:
                messagebox.showerror(ARABIC_TEXT["error"], message)
        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشل إنشاء القسم: {e}")

    def _delete_department(self, name: str):
        """حذف قسم"""
        try:
            if messagebox.askyesno(
                ARABIC_TEXT["confirm"],
                ARABIC_TEXT["confirm_delete"].format(name)
            ):
                success, message = self.data_manager.delete_department(name)
                if success:
                    self.refresh()
                    messagebox.showinfo(ARABIC_TEXT["success"], f"تم حذف القسم: {name}")
                else:
                    messagebox.showerror(ARABIC_TEXT["error"], message)
        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشل حذف القسم: {e}")

    def _create_dept_card(self, parent, name: str, stats: dict) -> ctk.CTkFrame:
        """إنشاء بطاقة قسم"""
        card = ctk.CTkFrame(parent, corner_radius=10)

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="x", padx=15, pady=12)

        # زر الحذف على اليسار
        del_btn = ctk.CTkButton(
            content,
            text=ARABIC_TEXT["delete"],
            width=80,
            height=34,
            font=ctk.CTkFont(size=11),
            fg_color=COLORS["danger"],
            hover_color=COLORS["danger_hover"],
            command=lambda: self._delete_department(name)
        )
        del_btn.pack(side="left")

        # المعلومات على اليمين
        info_frame = ctk.CTkFrame(content, fg_color="transparent")
        info_frame.pack(side="right", fill="x", expand=True)

        name_label = ctk.CTkLabel(
            info_frame,
            text=name,
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="e"
        )
        name_label.pack(anchor="e")

        stages = stats.get("stages", {})
        stage_counts = [f"{s}: {c}" for s, c in stages.items() if c > 0]
        total = stats.get("total_students", 0)

        stats_text = f"{ARABIC_TEXT['total']}: {total} {ARABIC_TEXT['students']}"
        if stage_counts:
            stats_text += "  |  " + " | ".join(stage_counts)

        stats_label = ctk.CTkLabel(
            info_frame,
            text=stats_text,
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_secondary"],
            anchor="e"
        )
        stats_label.pack(anchor="e", pady=(3, 0))

        return card

    def refresh(self):
        """تحديث قائمة الأقسام"""
        try:
            for widget in self.dept_widgets:
                widget.destroy()
            self.dept_widgets.clear()

            departments = self.data_manager.departments
            if not departments:
                empty_label = ctk.CTkLabel(
                    self.list_frame,
                    text=ARABIC_TEXT["no_departments"],
                    font=ctk.CTkFont(size=14),
                    text_color=COLORS["text_secondary"]
                )
                empty_label.pack(pady=50)
                self.dept_widgets.append(empty_label)
            else:
                for name in departments:
                    stats = self.data_manager.get_department_summary(name)
                    card = self._create_dept_card(self.list_frame, name, stats)
                    card.pack(fill="x", padx=5, pady=5)
                    self.dept_widgets.append(card)
        except Exception as e:
            print(f"Error refreshing departments: {e}")


# ==================== صفحة الطلاب ====================

class StudentsPage(ctk.CTkFrame):
    """صفحة إدارة الطلاب"""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.stage_textboxes: Dict[str, tuple] = {}

        self._create_header()
        self._create_import_section()
        self._create_manual_entry_section()
        self._create_list_section()

    def _create_header(self):
        """إنشاء رأس الصفحة"""
        self.title = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["students_title"],
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="e", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["students_subtitle"],
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="e", pady=(0, SCALE["padding_large"]))

    def _create_import_section(self):
        """إنشاء قسم الاستيراد"""
        self.import_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.import_card.pack(fill="x", pady=(0, SCALE["padding_medium"]))

        self.import_title = ctk.CTkLabel(
            self.import_card,
            text=ARABIC_TEXT["import_from_file"],
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.import_title.pack(anchor="e", padx=20, pady=(15, 10))

        # صف الاختيارات
        self.select_frame = ctk.CTkFrame(self.import_card, fg_color="transparent")
        self.select_frame.pack(fill="x", padx=20, pady=5)

        # المرحلة
        self.stage_var = ctk.StringVar(value=ARABIC_TEXT["stage_1"])
        self.stage_dropdown = ctk.CTkOptionMenu(
            self.select_frame,
            variable=self.stage_var,
            values=[ARABIC_TEXT["stage_1"], ARABIC_TEXT["stage_2"], ARABIC_TEXT["stage_3"]],
            width=140,
            height=38,
            font=ctk.CTkFont(size=12)
        )
        self.stage_dropdown.pack(side="right", padx=(10, 0))

        self.stage_label = ctk.CTkLabel(
            self.select_frame,
            text=ARABIC_TEXT["stage"] + ":",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.stage_label.pack(side="right", padx=(20, 5))

        # القسم
        self.dept_var = ctk.StringVar(value=ARABIC_TEXT["select_dept"])
        self.dept_dropdown = ctk.CTkOptionMenu(
            self.select_frame,
            variable=self.dept_var,
            values=[ARABIC_TEXT["select_dept"]],
            width=200,
            height=38,
            font=ctk.CTkFont(size=12),
            command=self._on_dept_change
        )
        self.dept_dropdown.pack(side="right", padx=(10, 0))

        self.dept_label = ctk.CTkLabel(
            self.select_frame,
            text=ARABIC_TEXT["department"] + ":",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.dept_label.pack(side="right", padx=(0, 5))

        # صف الأزرار
        self.btn_frame = ctk.CTkFrame(self.import_card, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=20, pady=(10, 15))

        self.clear_stage_btn = ctk.CTkButton(
            self.btn_frame,
            text=ARABIC_TEXT["clear_stage"],
            font=ctk.CTkFont(size=12),
            height=40,
            fg_color=COLORS["danger"],
            hover_color=COLORS["danger_hover"],
            command=self._clear_stage
        )
        self.clear_stage_btn.pack(side="left", padx=(0, 10))

        self.import_xlsx_btn = ctk.CTkButton(
            self.btn_frame,
            text=ARABIC_TEXT["import_excel"],
            font=ctk.CTkFont(size=12, weight="bold"),
            height=40,
            fg_color=COLORS["success"],
            hover_color=COLORS["success_hover"],
            command=lambda: self._import_file("xlsx")
        )
        self.import_xlsx_btn.pack(side="right", padx=(10, 0))

        self.import_docx_btn = ctk.CTkButton(
            self.btn_frame,
            text=ARABIC_TEXT["import_word"],
            font=ctk.CTkFont(size=12, weight="bold"),
            height=40,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            command=lambda: self._import_file("docx")
        )
        self.import_docx_btn.pack(side="right")

    def _create_manual_entry_section(self):
        """إنشاء قسم الإضافة اليدوية"""
        self.manual_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.manual_card.pack(fill="x", pady=(0, SCALE["padding_medium"]))

        self.manual_title = ctk.CTkLabel(
            self.manual_card,
            text=ARABIC_TEXT["add_manually"],
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.manual_title.pack(anchor="e", padx=20, pady=(15, 10))

        self.manual_frame = ctk.CTkFrame(self.manual_card, fg_color="transparent")
        self.manual_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.add_student_btn = ctk.CTkButton(
            self.manual_frame,
            text=ARABIC_TEXT["add_student"],
            font=ctk.CTkFont(size=12, weight="bold"),
            height=SCALE["button_height"],
            width=100,
            fg_color=COLORS["accent"],
            hover_color="#6D28D9",
            command=self._add_student_manually
        )
        self.add_student_btn.pack(side="left")

        self.student_entry = ctk.CTkEntry(
            self.manual_frame,
            placeholder_text=ARABIC_TEXT["student_name"],
            height=SCALE["entry_height"],
            font=ctk.CTkFont(size=SCALE["font_normal"]),
            justify="right"
        )
        self.student_entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
        self.student_entry.bind("<Return>", lambda e: self._add_student_manually())

    def _create_list_section(self):
        """إنشاء قسم القائمة"""
        self.list_label = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["student_lists"],
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.list_label.pack(anchor="e", pady=(10, SCALE["padding_medium"]))

        self.tabview = ctk.CTkTabview(self, corner_radius=SCALE["card_corner"])
        self.tabview.pack(fill="both", expand=True)

        self.stage_tabs = {}
        stages = [
            (ARABIC_TEXT["stage_1"], "1st Stage"),
            (ARABIC_TEXT["stage_2"], "2nd Stage"),
            (ARABIC_TEXT["stage_3"], "3rd Stage")
        ]

        for arabic_name, english_name in stages:
            tab = self.tabview.add(arabic_name)
            self.stage_tabs[english_name] = tab

            count_label = ctk.CTkLabel(
                tab,
                text=f"{ARABIC_TEXT['students_label']} 0",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=COLORS["primary_light"]
            )
            count_label.pack(anchor="e", padx=10, pady=(5, 0))

            textbox = ctk.CTkTextbox(tab, font=ctk.CTkFont(size=13))
            textbox.pack(fill="both", expand=True, padx=10, pady=(5, 10))
            self.stage_textboxes[english_name] = (textbox, count_label)

    def _get_stage_english(self, arabic_stage: str) -> str:
        """تحويل اسم المرحلة للإنجليزية"""
        stage_map = {
            ARABIC_TEXT["stage_1"]: "1st Stage",
            ARABIC_TEXT["stage_2"]: "2nd Stage",
            ARABIC_TEXT["stage_3"]: "3rd Stage",
        }
        return stage_map.get(arabic_stage, "1st Stage")

    def _on_dept_change(self, value):
        """معالجة تغيير القسم"""
        self.refresh()

    def _add_student_manually(self):
        """إضافة طالب يدوياً"""
        try:
            dept = self.dept_var.get()
            if dept == ARABIC_TEXT["select_dept"]:
                messagebox.showwarning(ARABIC_TEXT["warning"], ARABIC_TEXT["no_dept_selected"])
                return

            name = self.student_entry.get().strip()
            if not name:
                messagebox.showwarning(ARABIC_TEXT["warning"], "يرجى إدخال اسم الطالب")
                return

            stage = self._get_stage_english(self.stage_var.get())
            success = self.data_manager.add_student(dept, stage, name)

            if success:
                self.student_entry.delete(0, "end")
                self.refresh()
            else:
                messagebox.showinfo("معلومة", f"الطالب '{name}' موجود مسبقاً")
        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشلت الإضافة: {e}")

    def _import_file(self, file_type: str):
        """استيراد الطلاب من ملف"""
        try:
            dept = self.dept_var.get()
            if dept == ARABIC_TEXT["select_dept"]:
                messagebox.showwarning(ARABIC_TEXT["warning"], ARABIC_TEXT["no_dept_selected"])
                return

            stage = self._get_stage_english(self.stage_var.get())

            if file_type == "docx":
                filetypes = [("Word Documents", "*.docx"), ("All Files", "*.*")]
            else:
                filetypes = [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]

            file_path = filedialog.askopenfilename(
                title=f"اختر ملف {file_type.upper()}",
                filetypes=filetypes
            )

            if not file_path:
                return

            success, message, count = self.data_manager.import_from_file(file_path, dept, stage)

            if success:
                self.refresh()
                messagebox.showinfo(ARABIC_TEXT["import_success"], f"تم استيراد {count} طالب")
            else:
                messagebox.showerror(ARABIC_TEXT["error"], message)
        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشل الاستيراد: {e}")

    def _clear_stage(self):
        """مسح طلاب المرحلة"""
        try:
            dept = self.dept_var.get()
            if dept == ARABIC_TEXT["select_dept"]:
                messagebox.showwarning(ARABIC_TEXT["warning"], ARABIC_TEXT["no_dept_selected"])
                return

            stage = self._get_stage_english(self.stage_var.get())
            arabic_stage = self.stage_var.get()
            count = self.data_manager.get_student_count(dept, stage)

            if count == 0:
                messagebox.showinfo("معلومة", ARABIC_TEXT["no_students_stage"].format(arabic_stage))
                return

            if messagebox.askyesno(ARABIC_TEXT["confirm"], ARABIC_TEXT["clear_confirm"].format(count, arabic_stage)):
                self.data_manager.clear_stage(dept, stage)
                self.refresh()
                messagebox.showinfo(ARABIC_TEXT["success"], ARABIC_TEXT["cleared_students"].format(count, arabic_stage))
        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشل المسح: {e}")

    def refresh(self):
        """تحديث عرض الطلاب"""
        try:
            departments = self.data_manager.departments
            if departments:
                self.dept_dropdown.configure(values=departments)
                if self.dept_var.get() == ARABIC_TEXT["select_dept"] or self.dept_var.get() not in departments:
                    self.dept_var.set(departments[0])
            else:
                self.dept_dropdown.configure(values=[ARABIC_TEXT["select_dept"]])
                self.dept_var.set(ARABIC_TEXT["select_dept"])

            dept = self.dept_var.get()
            if dept != ARABIC_TEXT["select_dept"]:
                for stage, (textbox, count_label) in self.stage_textboxes.items():
                    students = self.data_manager.get_students(dept, stage)

                    textbox.configure(state="normal")
                    textbox.delete("1.0", "end")

                    if students:
                        for i, student in enumerate(students, 1):
                            textbox.insert("end", f"  {student}  .{i}\n")
                        count_label.configure(text=f"{ARABIC_TEXT['students_label']} {len(students)}")
                    else:
                        textbox.insert("1.0", ARABIC_TEXT["no_students"])
                        count_label.configure(text=f"{ARABIC_TEXT['students_label']} 0")

                    textbox.configure(state="disabled")
            else:
                for stage, (textbox, count_label) in self.stage_textboxes.items():
                    textbox.configure(state="normal")
                    textbox.delete("1.0", "end")
                    textbox.insert("1.0", ARABIC_TEXT["select_dept_first"])
                    textbox.configure(state="disabled")
                    count_label.configure(text=f"{ARABIC_TEXT['students_label']} 0")
        except Exception as e:
            print(f"Error refreshing students: {e}")


# ==================== صفحة توليد التوزيع ====================

class GeneratePage(ctk.CTkFrame):
    """صفحة توليد التوزيع"""

    def __init__(self, master, data_manager: DataManager, algorithm: ZigzagSeatingAlgorithm, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.algorithm = algorithm

        self._create_header()
        self._create_info_section()
        self._create_controls()
        self._create_results_section()

    def _create_header(self):
        """إنشاء رأس الصفحة"""
        self.title = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["generate_title"],
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="e", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["generate_subtitle"],
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="e", pady=(0, SCALE["padding_large"]))

    def _create_info_section(self):
        """إنشاء قسم المعلومات"""
        self.info_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.info_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.info_title = ctk.CTkLabel(
            self.info_card,
            text=ARABIC_TEXT["algorithm_info"],
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.info_title.pack(anchor="e", padx=20, pady=(15, 8))

        self.info_label = ctk.CTkLabel(
            self.info_card,
            text=ARABIC_TEXT["algorithm_desc"],
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"],
            justify="right"
        )
        self.info_label.pack(anchor="e", padx=20, pady=(0, 15))

    def _create_controls(self):
        """إنشاء عناصر التحكم"""
        self.control_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.control_frame.pack(fill="x", pady=(0, SCALE["padding_medium"]))

        self.generate_btn = ctk.CTkButton(
            self.control_frame,
            text=ARABIC_TEXT["generate_btn"],
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            width=180,
            fg_color=COLORS["warning"],
            hover_color=COLORS["warning_hover"],
            command=self._generate
        )
        self.generate_btn.pack(side="left")

        self.dept_var = ctk.StringVar(value=ARABIC_TEXT["select_dept"])
        self.dept_dropdown = ctk.CTkOptionMenu(
            self.control_frame,
            variable=self.dept_var,
            values=[ARABIC_TEXT["select_dept"]],
            width=220,
            height=45,
            font=ctk.CTkFont(size=12)
        )
        self.dept_dropdown.pack(side="right", padx=(10, 0))

        self.dept_label = ctk.CTkLabel(
            self.control_frame,
            text=ARABIC_TEXT["select_department"],
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.dept_label.pack(side="right")

    def _create_results_section(self):
        """إنشاء قسم النتائج"""
        self.results_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.results_card.pack(fill="both", expand=True)

        self.results_title = ctk.CTkLabel(
            self.results_card,
            text=ARABIC_TEXT["results_title"],
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.results_title.pack(anchor="e", padx=20, pady=(15, 10))

        self.results_text = ctk.CTkTextbox(
            self.results_card,
            font=ctk.CTkFont(size=11, family="Courier")
        )
        self.results_text.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self._show_initial_message()

        self.stats_label = ctk.CTkLabel(
            self.results_card,
            text="",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COLORS["success"]
        )
        self.stats_label.pack(anchor="e", padx=20, pady=(0, 15))

    def _show_initial_message(self):
        """عرض الرسالة الأولية"""
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")
        self.results_text.insert("1.0", ARABIC_TEXT["initial_message"] + "\n\n")
        self.results_text.insert("end", "الخوارزمية ستقوم بـ:\n")
        self.results_text.insert("end", "  - تبديل الطلاب بين المراحل بشكل صارم\n")
        self.results_text.insert("end", "  - ترتيب عشوائي داخل كل مرحلة\n")
        self.results_text.insert("end", "  - معالجة الأعداد الفردية بمقاعد فارغة\n")
        self.results_text.insert("end", "  - زيادة الأزواج المختلطة لمنع الغش\n")
        self.results_text.configure(state="disabled")

    def _generate(self):
        """توليد التوزيع"""
        try:
            dept_name = self.dept_var.get()
            if dept_name == ARABIC_TEXT["select_dept"]:
                messagebox.showwarning(ARABIC_TEXT["warning"], ARABIC_TEXT["no_dept_selected"])
                return

            is_valid, message = self.data_manager.validate_for_seating(dept_name)
            if not is_valid:
                messagebox.showerror(ARABIC_TEXT["validation_error"], message)
                return

            dept = self.data_manager.get_department(dept_name)
            if not dept:
                messagebox.showerror(ARABIC_TEXT["error"], "القسم غير موجود")
                return

            student_pools = dept.get_all_students()

            success, message, result = self.algorithm.generate(student_pools)

            if not success:
                messagebox.showerror(ARABIC_TEXT["generation_error"], message)
                return

            self._display_results(result)

            if result.warnings:
                messagebox.showwarning("ملاحظة", "\n".join(result.warnings))
            else:
                messagebox.showinfo(ARABIC_TEXT["success"], "تم توليد التوزيع بنجاح!")

        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشل التوليد: {e}")
            print(traceback.format_exc())

    def _display_results(self, result: SeatingResult):
        """عرض النتائج"""
        try:
            self.results_text.configure(state="normal")
            self.results_text.delete("1.0", "end")

            pattern = " <- ".join(result.stage_order) if result.stage_order else "N/A"
            self.results_text.insert("end", f"{ARABIC_TEXT['pattern']}: {pattern} (متكرر)\n")
            self.results_text.insert("end", "=" * 110 + "\n\n")

            header = f"{ARABIC_TEXT['student_b']:<30} {ARABIC_TEXT['seat_b']:<8} {ARABIC_TEXT['student_a']:<30} {ARABIC_TEXT['seat_a']:<8} {ARABIC_TEXT['row']:<5} {ARABIC_TEXT['column']:<8} {ARABIC_TEXT['desk']:<6}\n"
            self.results_text.insert("end", header)
            self.results_text.insert("end", "-" * 110 + "\n")

            for desk in result.desks:
                if desk.student_a:
                    if desk.student_a.is_empty:
                        student_a = ARABIC_TEXT["empty_seat"]
                    else:
                        name_a = desk.student_a.name[:20] + "..." if len(desk.student_a.name) > 20 else desk.student_a.name
                        student_a = f"({desk.student_a.stage[:3]}) {name_a}"
                else:
                    student_a = "-"

                if desk.student_b:
                    if desk.student_b.is_empty:
                        student_b = ARABIC_TEXT["empty_seat"]
                    else:
                        name_b = desk.student_b.name[:20] + "..." if len(desk.student_b.name) > 20 else desk.student_b.name
                        student_b = f"({desk.student_b.stage[:3]}) {name_b}"
                else:
                    student_b = "-"

                row = f"{student_b:<30} {desk.seat_b_number:<8} {student_a:<30} {desk.seat_a_number:<8} {desk.row:<5} {desk.column.value:<8} {desk.number:<6}\n"
                self.results_text.insert("end", row)

            self.results_text.configure(state="disabled")

            stats = result.stats
            empty_count = stats.get('empty_seats', 0)
            empty_text = f" | {ARABIC_TEXT['empty_seats']}: {empty_count}" if empty_count > 0 else ""

            self.stats_label.configure(
                text=f"تم بنجاح: {stats.get('total_desks', 0)} مقعد | "
                f"{stats.get('total_students', 0)} طالب | "
                f"{ARABIC_TEXT['cross_stage']}: {stats.get('cross_stage_pairs', 0)}{empty_text}"
            )
        except Exception as e:
            print(f"Error displaying results: {e}")

    def refresh(self):
        """تحديث الصفحة"""
        try:
            departments = self.data_manager.departments
            if departments:
                self.dept_dropdown.configure(values=departments)
                if self.dept_var.get() == ARABIC_TEXT["select_dept"] or self.dept_var.get() not in departments:
                    self.dept_var.set(departments[0])
            else:
                self.dept_dropdown.configure(values=[ARABIC_TEXT["select_dept"]])
                self.dept_var.set(ARABIC_TEXT["select_dept"])
        except Exception as e:
            print(f"Error refreshing generate page: {e}")


# ==================== صفحة التصدير ====================

class ExportPage(ctk.CTkFrame):
    """صفحة التصدير"""

    def __init__(
        self,
        master,
        data_manager: DataManager,
        algorithm: ZigzagSeatingAlgorithm,
        exporter: WordExporter,
        **kwargs
    ):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.algorithm = algorithm
        self.exporter = exporter

        self._create_header()
        self._create_options_section()
        self._create_export_button()
        self._create_preview_section()

    def _create_header(self):
        """إنشاء رأس الصفحة"""
        self.title = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["export_title"],
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="e", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text=ARABIC_TEXT["export_subtitle"],
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="e", pady=(0, SCALE["padding_large"]))

    def _create_options_section(self):
        """إنشاء قسم الخيارات"""
        self.options_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.options_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.options_title = ctk.CTkLabel(
            self.options_card,
            text=ARABIC_TEXT["export_options"],
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.options_title.pack(anchor="e", padx=20, pady=(15, 15))

        options_frame = ctk.CTkFrame(self.options_card, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=(0, 15))

        # القسم
        row1 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row1.pack(fill="x", pady=5)

        self.dept_var = ctk.StringVar(value=ARABIC_TEXT["select_dept"])
        self.dept_dropdown = ctk.CTkOptionMenu(
            row1,
            variable=self.dept_var,
            values=[ARABIC_TEXT["select_dept"]],
            width=250,
            height=38
        )
        self.dept_dropdown.pack(side="right", padx=(10, 0))

        ctk.CTkLabel(
            row1,
            text=ARABIC_TEXT["department"] + ":",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="right")

        # رقم القاعة
        row2 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row2.pack(fill="x", pady=5)

        self.room_entry = ctk.CTkEntry(row2, width=100, height=38, justify="right")
        self.room_entry.pack(side="right", padx=(10, 0))
        self.room_entry.insert(0, "1")

        ctk.CTkLabel(
            row2,
            text=ARABIC_TEXT["room_number"] + ":",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="right")

        # مدير الاعدادية
        row3 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row3.pack(fill="x", pady=5)

        self.director_entry = ctk.CTkEntry(row3, width=300, height=38, justify="right")
        self.director_entry.pack(side="right", padx=(10, 0))

        ctk.CTkLabel(
            row3,
            text=ARABIC_TEXT["director_name"] + ":",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="right")

        # عنوان الامتحان
        row4 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row4.pack(fill="x", pady=5)

        self.title_entry = ctk.CTkEntry(row4, width=350, height=38, justify="right")
        self.title_entry.pack(side="right", padx=(10, 0))
        self.title_entry.insert(0, "توزيع مقاعد الامتحان")

        ctk.CTkLabel(
            row4,
            text=ARABIC_TEXT["exam_title"] + ":",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="right")

        # التاريخ
        row5 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row5.pack(fill="x", pady=5)

        self.date_entry = ctk.CTkEntry(
            row5,
            width=200,
            height=38,
            placeholder_text=ARABIC_TEXT["date_placeholder"],
            justify="right"
        )
        self.date_entry.pack(side="right", padx=(10, 0))

        ctk.CTkLabel(
            row5,
            text=ARABIC_TEXT["exam_date"] + ":",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="right")

        # تضمين الخريطة
        row6 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row6.pack(fill="x", pady=5)

        self.map_var = ctk.BooleanVar(value=True)
        self.map_check = ctk.CTkCheckBox(
            row6,
            text=ARABIC_TEXT["include_map"],
            variable=self.map_var,
            font=ctk.CTkFont(size=12)
        )
        self.map_check.pack(side="right")

    def _create_export_button(self):
        """إنشاء زر التصدير"""
        self.export_btn = ctk.CTkButton(
            self,
            text=ARABIC_TEXT["export_btn"],
            font=ctk.CTkFont(size=15, weight="bold"),
            height=50,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            command=self._export
        )
        self.export_btn.pack(pady=SCALE["padding_large"])

    def _create_preview_section(self):
        """إنشاء قسم المعاينة"""
        self.preview_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.preview_card.pack(fill="both", expand=True)

        self.preview_title = ctk.CTkLabel(
            self.preview_card,
            text=ARABIC_TEXT["preview"],
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.preview_title.pack(anchor="e", padx=20, pady=(15, 10))

        self.preview_text = ctk.CTkTextbox(self.preview_card, font=ctk.CTkFont(size=12))
        self.preview_text.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        self._update_preview()
        self.preview_text.configure(state="disabled")

    def _update_preview(self):
        """تحديث المعاينة"""
        try:
            self.preview_text.configure(state="normal")
            self.preview_text.delete("1.0", "end")

            result = self.algorithm.result
            if not result or not result.desks:
                self.preview_text.insert("1.0", f"{ARABIC_TEXT['no_data']}\n\n")
                self.preview_text.insert("end", f"{ARABIC_TEXT['generate_first']}\n\n")
                self.preview_text.insert("end", "المستند سيتضمن:\n")
                self.preview_text.insert("end", "  - رأس الجامعة والقسم\n")
                self.preview_text.insert("end", "  - جدول التوزيع التفصيلي\n")
                self.preview_text.insert("end", "  - خريطة المقاعد المرئية\n")
                self.preview_text.insert("end", "  - الإحصائيات والتوقيع\n")
            else:
                self.preview_text.insert("1.0", f"{ARABIC_TEXT['ready_export']}\n\n")
                self.preview_text.insert("end", f"الملخص:\n")
                self.preview_text.insert("end", f"   - إجمالي المقاعد: {result.total_desks}\n")
                self.preview_text.insert("end", f"   - إجمالي الطلاب: {result.total_students}\n")
                self.preview_text.insert("end", f"   - الأزواج المختلطة: {result.cross_stage_count}\n")
                self.preview_text.insert("end", f"   - المقاعد الفارغة: {result.empty_seats_count}\n")

                if result.stage_order:
                    self.preview_text.insert("end", f"   - النمط: {' <- '.join(result.stage_order)}\n")

                if result.warnings:
                    self.preview_text.insert("end", f"\nملاحظات:\n")
                    for warning in result.warnings:
                        self.preview_text.insert("end", f"   - {warning}\n")

            self.preview_text.configure(state="disabled")
        except Exception as e:
            print(f"Error updating preview: {e}")

    def _export(self):
        """تصدير التقرير"""
        try:
            result = self.algorithm.result
            if not result or not result.desks:
                messagebox.showerror(ARABIC_TEXT["error"], ARABIC_TEXT["no_data"] + "\n" + ARABIC_TEXT["generate_first"])
                return

            dept_name = self.dept_var.get()
            if dept_name == ARABIC_TEXT["select_dept"]:
                messagebox.showwarning(ARABIC_TEXT["warning"], ARABIC_TEXT["no_dept_selected"])
                return

            exam_title = self.title_entry.get().strip() or "توزيع مقاعد الامتحان"
            exam_date = self.date_entry.get().strip() or None
            include_map = self.map_var.get()
            director_name = self.director_entry.get().strip()

            try:
                room_number = int(self.room_entry.get().strip() or "1")
            except ValueError:
                room_number = 1

            default_name = f"قاعة_{room_number}_{dept_name.replace(' ', '_')}_توزيع.docx"
            file_path = filedialog.asksaveasfilename(
                title="حفظ المستند",
                defaultextension=".docx",
                filetypes=[("Word Documents", "*.docx")],
                initialfile=default_name
            )

            if not file_path:
                return

            success, message = self.exporter.export(
                result=result,
                university_name=self.data_manager.university_name,
                department_name=dept_name,
                output_path=file_path,
                exam_title=exam_title,
                exam_date=exam_date,
                include_map=include_map,
                director_name=director_name,
                room_number=room_number
            )

            if success:
                messagebox.showinfo(ARABIC_TEXT["export_success"], message)

                if messagebox.askyesno(ARABIC_TEXT["open_file"], "هل تريد فتح الملف؟"):
                    try:
                        if os.name == 'nt':
                            os.startfile(file_path)
                        else:
                            os.system(f'xdg-open "{file_path}" &')
                    except Exception:
                        pass
            else:
                messagebox.showerror(ARABIC_TEXT["error"], message)

        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشل التصدير: {e}")
            print(traceback.format_exc())

    def refresh(self):
        """تحديث الصفحة"""
        try:
            departments = self.data_manager.departments
            if departments:
                self.dept_dropdown.configure(values=departments)
                if self.dept_var.get() == ARABIC_TEXT["select_dept"] or self.dept_var.get() not in departments:
                    self.dept_var.set(departments[0])
            else:
                self.dept_dropdown.configure(values=[ARABIC_TEXT["select_dept"]])
                self.dept_var.set(ARABIC_TEXT["select_dept"])

            self._update_preview()
        except Exception as e:
            print(f"Error refreshing export page: {e}")


# ==================== التطبيق الرئيسي ====================

class ExamSeatingApp(ctk.CTk):
    """نافذة التطبيق الرئيسية"""

    def __init__(self):
        super().__init__()

        # إعدادات النافذة
        self.title("نظام توزيع مقاعد الامتحانات - الإصدار 3.0")
        self.geometry("1350x900")
        self.minsize(1150, 750)

        # توسيط النافذة
        self._center_window()

        # تهيئة المكونات
        self.data_manager = DataManager()
        self.algorithm = ZigzagSeatingAlgorithm()
        self.exporter = WordExporter()

        # إعداد الشبكة
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # إنشاء الواجهة
        self._create_sidebar()
        self._create_content_area()
        self._create_pages()

        # عرض الصفحة الرئيسية
        self._show_page("dashboard")

    def _center_window(self):
        """توسيط النافذة"""
        try:
            self.update_idletasks()
            width = 1350
            height = 900
            x = (self.winfo_screenwidth() // 2) - (width // 2)
            y = (self.winfo_screenheight() // 2) - (height // 2)
            self.geometry(f'{width}x{height}+{x}+{y}')
        except Exception:
            pass

    def _create_sidebar(self):
        """إنشاء القائمة الجانبية"""
        self.sidebar = Sidebar(
            self,
            nav_callback=self._navigate,
            clear_callback=self._clear_all_data
        )
        self.sidebar.grid(row=0, column=0, sticky="nsw")

    def _create_content_area(self):
        """إنشاء منطقة المحتوى"""
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

    def _create_pages(self):
        """إنشاء جميع الصفحات"""
        self.pages: Dict[str, ctk.CTkFrame] = {}

        self.pages["dashboard"] = DashboardPage(
            self.content_frame,
            self.data_manager,
            nav_callback=self._navigate
        )
        self.pages["departments"] = DepartmentsPage(self.content_frame, self.data_manager)
        self.pages["students"] = StudentsPage(self.content_frame, self.data_manager)
        self.pages["generate"] = GeneratePage(self.content_frame, self.data_manager, self.algorithm)
        self.pages["export"] = ExportPage(self.content_frame, self.data_manager, self.algorithm, self.exporter)

        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

    def _show_page(self, key: str):
        """عرض صفحة معينة"""
        try:
            page = self.pages.get(key)
            if page:
                page.tkraise()
                if hasattr(page, 'refresh'):
                    page.refresh()
        except Exception as e:
            print(f"Error showing page {key}: {e}")

    def _navigate(self, key: str):
        """معالجة التنقل"""
        self._show_page(key)
        self.sidebar._highlight(key)

    def _clear_all_data(self):
        """مسح جميع البيانات"""
        try:
            if messagebox.askyesno(
                ARABIC_TEXT["confirm"],
                ARABIC_TEXT["confirm_clear"]
            ):
                self.data_manager = DataManager()
                self.algorithm.clear()

                for page in self.pages.values():
                    page.destroy()
                self.pages.clear()

                self._create_pages()
                self._show_page("dashboard")

                messagebox.showinfo(ARABIC_TEXT["success"], ARABIC_TEXT["data_cleared"])
        except Exception as e:
            messagebox.showerror(ARABIC_TEXT["error"], f"فشل المسح: {e}")


# ==================== نقطة البدء ====================

def main():
    """نقطة بدء التطبيق"""
    try:
        app = ExamSeatingApp()
        app.mainloop()
    except Exception as e:
        print(f"Application error: {e}")
        print(traceback.format_exc())
        messagebox.showerror("خطأ فادح", f"فشل تشغيل التطبيق:\n{e}")


if __name__ == "__main__":
    main()
