#!/usr/bin/env python3
"""
University Exam Seating System
A modern desktop application for anti-cheating exam seating management.

Features:
- Multi-department support
- Import students from .docx and .xlsx files
- Zigzag anti-cheating seating algorithm
- Professional Word document export with visual seating maps
- Modern dark-themed UI with CustomTkinter
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Optional, Callable
import os

from data_manager import DataManager
from seating_algorithm import ZigzagSeatingAlgorithm, SeatingResult
from word_exporter import WordExporter


# ==================== Theme Configuration ====================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Color Scheme
COLORS = {
    "primary": "#3B82F6",
    "primary_hover": "#2563EB",
    "secondary": "#6B7280",
    "success": "#10B981",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "bg_dark": "#1F2937",
    "bg_card": "#374151",
    "text": "#F9FAFB",
    "text_secondary": "#9CA3AF"
}


# ==================== Sidebar Navigation ====================

class Sidebar(ctk.CTkFrame):
    """Modern sidebar navigation component."""

    def __init__(self, master, nav_callback: Callable, **kwargs):
        super().__init__(master, width=220, corner_radius=0, **kwargs)

        self.nav_callback = nav_callback
        self.buttons = {}
        self.current_page = "dashboard"

        # Prevent sidebar from shrinking
        self.grid_propagate(False)

        # App Header
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=15, pady=(20, 5))

        self.logo_label = ctk.CTkLabel(
            self.header_frame,
            text="🎓",
            font=ctk.CTkFont(size=36)
        )
        self.logo_label.pack()

        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text="Exam Seating",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.title_label.pack()

        self.subtitle_label = ctk.CTkLabel(
            self.header_frame,
            text="Management System",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle_label.pack()

        # Divider
        self.divider = ctk.CTkFrame(self, height=2, fg_color=COLORS["secondary"])
        self.divider.pack(fill="x", padx=15, pady=20)

        # Navigation Buttons
        nav_items = [
            ("dashboard", "📊", "Dashboard"),
            ("departments", "🏛️", "Departments"),
            ("students", "👥", "Students"),
            ("generate", "🎲", "Generate Seats"),
            ("export", "📄", "Export Report"),
        ]

        for key, icon, text in nav_items:
            btn = ctk.CTkButton(
                self,
                text=f"  {icon}  {text}",
                font=ctk.CTkFont(size=14),
                height=45,
                corner_radius=10,
                fg_color="transparent",
                text_color=COLORS["text"],
                hover_color=COLORS["bg_card"],
                anchor="w",
                command=lambda k=key: self._navigate(k)
            )
            btn.pack(fill="x", padx=10, pady=3)
            self.buttons[key] = btn

        # Set initial highlight
        self._highlight("dashboard")

        # Spacer
        self.spacer = ctk.CTkFrame(self, fg_color="transparent")
        self.spacer.pack(fill="both", expand=True)

        # Theme Toggle
        self.theme_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.theme_frame.pack(fill="x", padx=15, pady=10)

        self.theme_label = ctk.CTkLabel(
            self.theme_frame,
            text="Theme",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"]
        )
        self.theme_label.pack(anchor="w")

        self.theme_switch = ctk.CTkSwitch(
            self.theme_frame,
            text="Dark Mode",
            command=self._toggle_theme,
            onvalue="dark",
            offvalue="light"
        )
        self.theme_switch.pack(anchor="w", pady=5)
        self.theme_switch.select()

        # Version
        self.version_label = ctk.CTkLabel(
            self,
            text="v2.0.0",
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_secondary"]
        )
        self.version_label.pack(pady=(0, 15))

    def _navigate(self, key: str):
        """Handle navigation click."""
        self.current_page = key
        self._highlight(key)
        self.nav_callback(key)

    def _highlight(self, key: str):
        """Highlight the active navigation button."""
        for btn_key, btn in self.buttons.items():
            if btn_key == key:
                btn.configure(fg_color=COLORS["primary"])
            else:
                btn.configure(fg_color="transparent")

    def _toggle_theme(self):
        """Toggle between dark and light theme."""
        mode = self.theme_switch.get()
        ctk.set_appearance_mode(mode)


# ==================== Dashboard Page ====================

class DashboardPage(ctk.CTkFrame):
    """Main dashboard with overview statistics."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager

        # Page Title
        self.title = ctk.CTkLabel(
            self,
            text="Dashboard",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Overview of your exam seating system",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, 20))

        # University Name Card
        self.uni_card = ctk.CTkFrame(self, corner_radius=15)
        self.uni_card.pack(fill="x", pady=(0, 20))

        self.uni_label = ctk.CTkLabel(
            self.uni_card,
            text="University Name",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"]
        )
        self.uni_label.pack(anchor="w", padx=20, pady=(15, 5))

        self.uni_entry = ctk.CTkEntry(
            self.uni_card,
            placeholder_text="Enter university name...",
            height=45,
            font=ctk.CTkFont(size=14)
        )
        self.uni_entry.pack(fill="x", padx=20, pady=(0, 15))
        self.uni_entry.insert(0, "University Name")
        self.uni_entry.bind("<KeyRelease>", self._on_uni_change)

        # Stats Grid
        self.stats_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.stats_frame.pack(fill="x", pady=10)

        self.stats_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self.stat_cards = {}
        stats_config = [
            ("departments", "🏛️", "Departments", "0"),
            ("students", "👥", "Total Students", "0"),
            ("stages", "📚", "Active Stages", "0")
        ]

        for i, (key, icon, label, value) in enumerate(stats_config):
            card = self._create_stat_card(self.stats_frame, icon, label, value)
            card.grid(row=0, column=i, padx=5, pady=5, sticky="nsew")
            self.stat_cards[key] = card

        # Quick Actions
        self.actions_label = ctk.CTkLabel(
            self,
            text="Quick Actions",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.actions_label.pack(anchor="w", pady=(30, 15))

        self.actions_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.actions_frame.pack(fill="x")

        actions = [
            ("➕ Add Department", COLORS["primary"], self._quick_add_dept),
            ("📥 Import Students", COLORS["success"], self._quick_import),
            ("🎲 Generate Seating", COLORS["warning"], self._quick_generate)
        ]

        for text, color, cmd in actions:
            btn = ctk.CTkButton(
                self.actions_frame,
                text=text,
                font=ctk.CTkFont(size=13),
                height=40,
                fg_color=color,
                hover_color=self._darken_color(color),
                command=cmd
            )
            btn.pack(side="left", padx=(0, 10))

        # Recent Activity
        self.activity_label = ctk.CTkLabel(
            self,
            text="System Status",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.activity_label.pack(anchor="w", pady=(30, 15))

        self.activity_card = ctk.CTkFrame(self, corner_radius=15)
        self.activity_card.pack(fill="both", expand=True)

        self.activity_text = ctk.CTkTextbox(
            self.activity_card,
            font=ctk.CTkFont(size=12),
            fg_color="transparent"
        )
        self.activity_text.pack(fill="both", expand=True, padx=15, pady=15)
        self.activity_text.insert("1.0", "Welcome to Exam Seating Management System!\n\n")
        self.activity_text.insert("end", "• Create departments for your university\n")
        self.activity_text.insert("end", "• Import student lists from Word or Excel files\n")
        self.activity_text.insert("end", "• Generate anti-cheating seating arrangements\n")
        self.activity_text.insert("end", "• Export professional reports with visual maps\n")
        self.activity_text.configure(state="disabled")

    def _create_stat_card(self, parent, icon: str, label: str, value: str) -> ctk.CTkFrame:
        """Create a statistics card."""
        card = ctk.CTkFrame(parent, corner_radius=15, height=120)
        card.pack_propagate(False)

        icon_label = ctk.CTkLabel(card, text=icon, font=ctk.CTkFont(size=30))
        icon_label.pack(pady=(15, 5))

        value_label = ctk.CTkLabel(
            card,
            text=value,
            font=ctk.CTkFont(size=24, weight="bold")
        )
        value_label.pack()

        text_label = ctk.CTkLabel(
            card,
            text=label,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"]
        )
        text_label.pack()

        # Store reference to value label for updates
        card.value_label = value_label

        return card

    def _darken_color(self, hex_color: str) -> str:
        """Darken a hex color for hover effect."""
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        darkened = tuple(max(0, int(c * 0.8)) for c in rgb)
        return f"#{darkened[0]:02x}{darkened[1]:02x}{darkened[2]:02x}"

    def _on_uni_change(self, event):
        """Update university name."""
        self.data_manager.university_name = self.uni_entry.get()

    def _quick_add_dept(self):
        """Quick action to add department."""
        # This will be connected to navigation callback
        pass

    def _quick_import(self):
        """Quick action to import students."""
        pass

    def _quick_generate(self):
        """Quick action to generate seating."""
        pass

    def refresh(self):
        """Refresh dashboard statistics."""
        # Update university name
        if self.uni_entry.get() != self.data_manager.university_name:
            self.uni_entry.delete(0, "end")
            self.uni_entry.insert(0, self.data_manager.university_name)

        # Update stats
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


# ==================== Departments Page ====================

class DepartmentsPage(ctk.CTkFrame):
    """Department management page."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager

        # Page Title
        self.title = ctk.CTkLabel(
            self,
            text="Departments",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Manage university departments",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, 20))

        # Add Department Card
        self.add_card = ctk.CTkFrame(self, corner_radius=15)
        self.add_card.pack(fill="x", pady=(0, 20))

        self.add_label = ctk.CTkLabel(
            self.add_card,
            text="Create New Department",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.add_label.pack(anchor="w", padx=20, pady=(15, 10))

        self.add_frame = ctk.CTkFrame(self.add_card, fg_color="transparent")
        self.add_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.dept_entry = ctk.CTkEntry(
            self.add_frame,
            placeholder_text="Enter department name (e.g., Computer Science)",
            height=45,
            font=ctk.CTkFont(size=14)
        )
        self.dept_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.add_btn = ctk.CTkButton(
            self.add_frame,
            text="➕ Create",
            font=ctk.CTkFont(size=14),
            height=45,
            width=120,
            fg_color=COLORS["success"],
            hover_color="#059669",
            command=self._create_department
        )
        self.add_btn.pack(side="right")

        # Departments List
        self.list_label = ctk.CTkLabel(
            self,
            text="Existing Departments",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.list_label.pack(anchor="w", pady=(10, 15))

        self.list_frame = ctk.CTkScrollableFrame(self, corner_radius=15)
        self.list_frame.pack(fill="both", expand=True)

        self.dept_widgets = []

    def _create_department(self):
        """Create a new department."""
        name = self.dept_entry.get().strip()
        if not name:
            messagebox.showwarning("Warning", "Please enter a department name.")
            return

        success, message = self.data_manager.create_department(name)
        if success:
            self.dept_entry.delete(0, "end")
            self.refresh()
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)

    def _delete_department(self, name: str):
        """Delete a department."""
        if messagebox.askyesno("Confirm", f"Delete department '{name}'?\nThis will remove all student data."):
            success, message = self.data_manager.delete_department(name)
            if success:
                self.refresh()
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)

    def _create_dept_card(self, parent, name: str, stats: dict) -> ctk.CTkFrame:
        """Create a department card widget."""
        card = ctk.CTkFrame(parent, corner_radius=10, height=80)

        # Department name
        name_label = ctk.CTkLabel(
            card,
            text=f"🏛️ {name}",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        name_label.pack(anchor="w", padx=15, pady=(12, 5))

        # Stats
        stage_info = " | ".join([
            f"{stage}: {count}"
            for stage, count in stats.get("stages", {}).items()
        ])
        total = stats.get("total_students", 0)

        stats_label = ctk.CTkLabel(
            card,
            text=f"Total: {total} students  •  {stage_info}",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_secondary"]
        )
        stats_label.pack(anchor="w", padx=15)

        # Delete button
        del_btn = ctk.CTkButton(
            card,
            text="🗑️",
            width=35,
            height=35,
            fg_color=COLORS["danger"],
            hover_color="#DC2626",
            command=lambda: self._delete_department(name)
        )
        del_btn.place(relx=0.95, rely=0.5, anchor="e")

        return card

    def refresh(self):
        """Refresh the departments list."""
        # Clear existing widgets
        for widget in self.dept_widgets:
            widget.destroy()
        self.dept_widgets.clear()

        # Create new cards
        departments = self.data_manager.departments
        if not departments:
            empty_label = ctk.CTkLabel(
                self.list_frame,
                text="No departments created yet.\nCreate your first department above.",
                font=ctk.CTkFont(size=14),
                text_color=COLORS["text_secondary"]
            )
            empty_label.pack(pady=40)
            self.dept_widgets.append(empty_label)
        else:
            for name in departments:
                stats = self.data_manager.get_department_summary(name)
                card = self._create_dept_card(self.list_frame, name, stats)
                card.pack(fill="x", padx=5, pady=5)
                self.dept_widgets.append(card)


# ==================== Students Page ====================

class StudentsPage(ctk.CTkFrame):
    """Student management and import page."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager

        # Page Title
        self.title = ctk.CTkLabel(
            self,
            text="Students",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Import and manage student lists",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, 20))

        # Import Section
        self.import_card = ctk.CTkFrame(self, corner_radius=15)
        self.import_card.pack(fill="x", pady=(0, 20))

        self.import_title = ctk.CTkLabel(
            self.import_card,
            text="Import Students",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.import_title.pack(anchor="w", padx=20, pady=(15, 10))

        # Department Selection
        self.select_frame = ctk.CTkFrame(self.import_card, fg_color="transparent")
        self.select_frame.pack(fill="x", padx=20, pady=5)

        self.dept_label = ctk.CTkLabel(
            self.select_frame,
            text="Department:",
            font=ctk.CTkFont(size=12)
        )
        self.dept_label.pack(side="left", padx=(0, 10))

        self.dept_var = ctk.StringVar(value="Select Department")
        self.dept_dropdown = ctk.CTkOptionMenu(
            self.select_frame,
            variable=self.dept_var,
            values=["Select Department"],
            width=200,
            height=35
        )
        self.dept_dropdown.pack(side="left", padx=(0, 20))

        self.stage_label = ctk.CTkLabel(
            self.select_frame,
            text="Stage:",
            font=ctk.CTkFont(size=12)
        )
        self.stage_label.pack(side="left", padx=(0, 10))

        self.stage_var = ctk.StringVar(value="1st Stage")
        self.stage_dropdown = ctk.CTkOptionMenu(
            self.select_frame,
            variable=self.stage_var,
            values=["1st Stage", "2nd Stage", "3rd Stage"],
            width=150,
            height=35
        )
        self.stage_dropdown.pack(side="left")

        # Import Buttons
        self.btn_frame = ctk.CTkFrame(self.import_card, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=20, pady=(10, 15))

        self.import_docx_btn = ctk.CTkButton(
            self.btn_frame,
            text="📄 Import from Word (.docx)",
            font=ctk.CTkFont(size=13),
            height=40,
            fg_color=COLORS["primary"],
            command=lambda: self._import_file("docx")
        )
        self.import_docx_btn.pack(side="left", padx=(0, 10))

        self.import_xlsx_btn = ctk.CTkButton(
            self.btn_frame,
            text="📊 Import from Excel (.xlsx)",
            font=ctk.CTkFont(size=13),
            height=40,
            fg_color=COLORS["success"],
            command=lambda: self._import_file("xlsx")
        )
        self.import_xlsx_btn.pack(side="left")

        # Student List View
        self.list_label = ctk.CTkLabel(
            self,
            text="Student Lists by Stage",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.list_label.pack(anchor="w", pady=(10, 15))

        # Tabview for stages
        self.tabview = ctk.CTkTabview(self, corner_radius=15)
        self.tabview.pack(fill="both", expand=True)

        self.stage_tabs = {}
        self.stage_textboxes = {}

        for stage in ["1st Stage", "2nd Stage", "3rd Stage"]:
            tab = self.tabview.add(stage)
            self.stage_tabs[stage] = tab

            textbox = ctk.CTkTextbox(tab, font=ctk.CTkFont(size=12))
            textbox.pack(fill="both", expand=True, padx=10, pady=10)
            self.stage_textboxes[stage] = textbox

        # Clear button
        self.clear_btn = ctk.CTkButton(
            self,
            text="🗑️ Clear Selected Stage",
            font=ctk.CTkFont(size=12),
            height=35,
            fg_color=COLORS["danger"],
            hover_color="#DC2626",
            command=self._clear_stage
        )
        self.clear_btn.pack(anchor="e", pady=(10, 0))

    def _import_file(self, file_type: str):
        """Import students from file."""
        dept = self.dept_var.get()
        if dept == "Select Department":
            messagebox.showwarning("Warning", "Please select a department first.")
            return

        stage = self.stage_var.get()

        if file_type == "docx":
            filetypes = [("Word Documents", "*.docx")]
        else:
            filetypes = [("Excel Files", "*.xlsx *.xls")]

        file_path = filedialog.askopenfilename(
            title=f"Select {file_type.upper()} File",
            filetypes=filetypes
        )

        if not file_path:
            return

        success, message, count = self.data_manager.import_from_file(file_path, dept, stage)

        if success:
            self.refresh()
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Import Error", message)

    def _clear_stage(self):
        """Clear students from selected stage."""
        dept = self.dept_var.get()
        if dept == "Select Department":
            messagebox.showwarning("Warning", "Please select a department.")
            return

        stage = self.stage_var.get()
        count = self.data_manager.get_student_count(dept, stage)

        if count == 0:
            messagebox.showinfo("Info", f"No students in {stage} to clear.")
            return

        if messagebox.askyesno("Confirm", f"Clear all {count} students from {stage}?"):
            self.data_manager.clear_stage(dept, stage)
            self.refresh()
            messagebox.showinfo("Success", f"Cleared {count} students from {stage}.")

    def refresh(self):
        """Refresh the students view."""
        # Update department dropdown
        departments = self.data_manager.departments
        if departments:
            self.dept_dropdown.configure(values=departments)
            if self.dept_var.get() == "Select Department":
                self.dept_var.set(departments[0])
        else:
            self.dept_dropdown.configure(values=["Select Department"])
            self.dept_var.set("Select Department")

        # Update student lists
        dept = self.dept_var.get()
        if dept != "Select Department":
            for stage, textbox in self.stage_textboxes.items():
                students = self.data_manager.get_students(dept, stage)

                textbox.configure(state="normal")
                textbox.delete("1.0", "end")

                if students:
                    for i, student in enumerate(students, 1):
                        textbox.insert("end", f"{i}. {student}\n")
                else:
                    textbox.insert("1.0", "No students imported for this stage.")

                textbox.configure(state="disabled")


# ==================== Generate Seating Page ====================

class GeneratePage(ctk.CTkFrame):
    """Seating generation page."""

    def __init__(self, master, data_manager: DataManager, algorithm: ZigzagSeatingAlgorithm, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.algorithm = algorithm
        self.current_result: Optional[SeatingResult] = None

        # Page Title
        self.title = ctk.CTkLabel(
            self,
            text="Generate Seating",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Create anti-cheating seating arrangements",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, 20))

        # Algorithm Info Card
        self.info_card = ctk.CTkFrame(self, corner_radius=15)
        self.info_card.pack(fill="x", pady=(0, 20))

        self.info_title = ctk.CTkLabel(
            self.info_card,
            text="🛡️ Anti-Cheating Algorithm",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.info_title.pack(anchor="w", padx=20, pady=(15, 10))

        info_text = (
            "• Students from DIFFERENT stages are paired at each desk\n"
            "• Students within each stage are RANDOMIZED\n"
            "• Desks arranged in ZIGZAG pattern across 3 columns\n"
            "• Handles unbalanced stage counts gracefully"
        )
        self.info_label = ctk.CTkLabel(
            self.info_card,
            text=info_text,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"],
            justify="left"
        )
        self.info_label.pack(anchor="w", padx=20, pady=(0, 15))

        # Department Selection
        self.select_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.select_frame.pack(fill="x", pady=(0, 15))

        self.dept_label = ctk.CTkLabel(
            self.select_frame,
            text="Select Department:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.dept_label.pack(side="left", padx=(0, 15))

        self.dept_var = ctk.StringVar(value="Select Department")
        self.dept_dropdown = ctk.CTkOptionMenu(
            self.select_frame,
            variable=self.dept_var,
            values=["Select Department"],
            width=250,
            height=40
        )
        self.dept_dropdown.pack(side="left", padx=(0, 20))

        self.generate_btn = ctk.CTkButton(
            self.select_frame,
            text="🎲 Generate Seating",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            width=180,
            fg_color=COLORS["warning"],
            hover_color="#D97706",
            command=self._generate
        )
        self.generate_btn.pack(side="left")

        # Results Area
        self.results_card = ctk.CTkFrame(self, corner_radius=15)
        self.results_card.pack(fill="both", expand=True)

        self.results_title = ctk.CTkLabel(
            self.results_card,
            text="Generated Seating Arrangement",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.results_title.pack(anchor="w", padx=20, pady=(15, 10))

        self.results_text = ctk.CTkTextbox(
            self.results_card,
            font=ctk.CTkFont(size=11, family="Courier")
        )
        self.results_text.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        self.results_text.insert("1.0", "Select a department and click 'Generate Seating' to begin...")
        self.results_text.configure(state="disabled")

        # Stats bar
        self.stats_label = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"]
        )
        self.stats_label.pack(anchor="w", pady=(10, 0))

    def _generate(self):
        """Generate seating arrangement."""
        dept_name = self.dept_var.get()
        if dept_name == "Select Department":
            messagebox.showwarning("Warning", "Please select a department.")
            return

        # Validate
        is_valid, message = self.data_manager.validate_for_seating(dept_name)
        if not is_valid:
            messagebox.showerror("Validation Error", message)
            return

        # Get student data
        dept = self.data_manager.get_department(dept_name)
        if not dept:
            messagebox.showerror("Error", "Department not found.")
            return

        student_pools = dept.get_all_students()

        # Generate
        success, message, result = self.algorithm.generate(student_pools)

        if not success:
            messagebox.showerror("Generation Error", message)
            return

        self.current_result = result
        self._display_results(result)

        # Show warnings if any
        if result.warnings:
            messagebox.showwarning("Warning", "\n".join(result.warnings))
        else:
            messagebox.showinfo("Success", message)

    def _display_results(self, result: SeatingResult):
        """Display the generated seating."""
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")

        # Header
        header = f"{'Desk':<6} {'Column':<8} {'Row':<5} {'Student A':<28} {'Student B':<28} {'Status':<12}\n"
        self.results_text.insert("end", header)
        self.results_text.insert("end", "=" * 100 + "\n")

        # Data
        for desk in result.desks:
            name_a = desk.student_a.name if desk.student_a else "-"
            stage_a = f"({desk.student_a.stage})" if desk.student_a else ""
            student_a = f"{name_a[:20]} {stage_a}"

            name_b = desk.student_b.name if desk.student_b else "-"
            stage_b = f"({desk.student_b.stage})" if desk.student_b else ""
            student_b = f"{name_b[:20]} {stage_b}"

            status = "✓ Cross" if desk.is_cross_stage else ("⚠ Same" if desk.is_full else "Single")

            row = f"{desk.number:<6} {desk.column.value:<8} {desk.row:<5} {student_a:<28} {student_b:<28} {status:<12}\n"
            self.results_text.insert("end", row)

        self.results_text.configure(state="disabled")

        # Update stats
        stats = result.stats
        self.stats_label.configure(
            text=f"✅ {stats['total_desks']} desks | "
            f"{stats['total_students']} students | "
            f"{stats['cross_stage_pairs']} cross-stage | "
            f"{stats['same_stage_pairs']} same-stage"
        )

    def refresh(self):
        """Refresh the page."""
        departments = self.data_manager.departments
        if departments:
            self.dept_dropdown.configure(values=departments)
            if self.dept_var.get() == "Select Department":
                self.dept_var.set(departments[0])
        else:
            self.dept_dropdown.configure(values=["Select Department"])
            self.dept_var.set("Select Department")


# ==================== Export Page ====================

class ExportPage(ctk.CTkFrame):
    """Export report page."""

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

        # Page Title
        self.title = ctk.CTkLabel(
            self,
            text="Export Report",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Generate professional Word documents",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, 20))

        # Export Options Card
        self.options_card = ctk.CTkFrame(self, corner_radius=15)
        self.options_card.pack(fill="x", pady=(0, 20))

        self.options_title = ctk.CTkLabel(
            self.options_card,
            text="Export Options",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.options_title.pack(anchor="w", padx=20, pady=(15, 15))

        # Department
        self.dept_frame = ctk.CTkFrame(self.options_card, fg_color="transparent")
        self.dept_frame.pack(fill="x", padx=20, pady=5)

        self.dept_label = ctk.CTkLabel(
            self.dept_frame,
            text="Department:",
            font=ctk.CTkFont(size=12),
            width=100
        )
        self.dept_label.pack(side="left")

        self.dept_var = ctk.StringVar(value="Select Department")
        self.dept_dropdown = ctk.CTkOptionMenu(
            self.dept_frame,
            variable=self.dept_var,
            values=["Select Department"],
            width=250
        )
        self.dept_dropdown.pack(side="left")

        # Exam Title
        self.title_frame = ctk.CTkFrame(self.options_card, fg_color="transparent")
        self.title_frame.pack(fill="x", padx=20, pady=5)

        self.title_label = ctk.CTkLabel(
            self.title_frame,
            text="Exam Title:",
            font=ctk.CTkFont(size=12),
            width=100
        )
        self.title_label.pack(side="left")

        self.title_entry = ctk.CTkEntry(
            self.title_frame,
            placeholder_text="e.g., Final Exam Seating Arrangement",
            width=400
        )
        self.title_entry.pack(side="left")
        self.title_entry.insert(0, "Exam Seating Arrangement")

        # Date
        self.date_frame = ctk.CTkFrame(self.options_card, fg_color="transparent")
        self.date_frame.pack(fill="x", padx=20, pady=5)

        self.date_label = ctk.CTkLabel(
            self.date_frame,
            text="Exam Date:",
            font=ctk.CTkFont(size=12),
            width=100
        )
        self.date_label.pack(side="left")

        self.date_entry = ctk.CTkEntry(
            self.date_frame,
            placeholder_text="Leave empty for today's date",
            width=250
        )
        self.date_entry.pack(side="left")

        # Include Map Option
        self.map_frame = ctk.CTkFrame(self.options_card, fg_color="transparent")
        self.map_frame.pack(fill="x", padx=20, pady=(10, 15))

        self.map_var = ctk.BooleanVar(value=True)
        self.map_check = ctk.CTkCheckBox(
            self.map_frame,
            text="Include Visual Seating Map",
            variable=self.map_var,
            font=ctk.CTkFont(size=12)
        )
        self.map_check.pack(anchor="w")

        # Export Button
        self.export_btn = ctk.CTkButton(
            self,
            text="📄 Export to Word Document",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            fg_color=COLORS["primary"],
            command=self._export
        )
        self.export_btn.pack(pady=20)

        # Preview/Status Card
        self.status_card = ctk.CTkFrame(self, corner_radius=15)
        self.status_card.pack(fill="both", expand=True)

        self.status_title = ctk.CTkLabel(
            self.status_card,
            text="Export Preview",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.status_title.pack(anchor="w", padx=20, pady=(15, 10))

        self.status_text = ctk.CTkTextbox(
            self.status_card,
            font=ctk.CTkFont(size=12)
        )
        self.status_text.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        self._update_preview()
        self.status_text.configure(state="disabled")

    def _update_preview(self):
        """Update the preview text."""
        self.status_text.configure(state="normal")
        self.status_text.delete("1.0", "end")

        result = self.algorithm.result
        if not result or not result.desks:
            self.status_text.insert("1.0", "No seating arrangement generated yet.\n\n")
            self.status_text.insert("end", "Please go to 'Generate Seats' page first to create\n")
            self.status_text.insert("end", "a seating arrangement before exporting.")
        else:
            self.status_text.insert("1.0", "✅ Seating arrangement ready for export!\n\n")
            self.status_text.insert("end", f"• Total Desks: {result.total_desks}\n")
            self.status_text.insert("end", f"• Total Students: {result.total_students}\n")
            self.status_text.insert("end", f"• Cross-Stage Pairs: {result.cross_stage_count}\n")
            self.status_text.insert("end", f"• Same-Stage Pairs: {result.same_stage_count}\n\n")

            if result.warnings:
                self.status_text.insert("end", "⚠️ Warnings:\n")
                for warning in result.warnings:
                    self.status_text.insert("end", f"  • {warning}\n")

        self.status_text.configure(state="disabled")

    def _export(self):
        """Export the report."""
        result = self.algorithm.result
        if not result or not result.desks:
            messagebox.showerror("Error", "No seating arrangement to export.\nPlease generate seating first.")
            return

        dept_name = self.dept_var.get()
        if dept_name == "Select Department":
            messagebox.showwarning("Warning", "Please select a department.")
            return

        # Get options
        exam_title = self.title_entry.get() or "Exam Seating Arrangement"
        exam_date = self.date_entry.get() or None
        include_map = self.map_var.get()

        # Get save path
        file_path = filedialog.asksaveasfilename(
            title="Save Word Document",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            initialfilename=f"{dept_name.replace(' ', '_')}_seating.docx"
        )

        if not file_path:
            return

        # Export
        success, message = self.exporter.export(
            result=result,
            university_name=self.data_manager.university_name,
            department_name=dept_name,
            output_path=file_path,
            exam_title=exam_title,
            exam_date=exam_date,
            include_map=include_map
        )

        if success:
            messagebox.showinfo("Success", message)

            if messagebox.askyesno("Open File", "Would you like to open the exported file?"):
                try:
                    if os.name == 'nt':
                        os.startfile(file_path)
                    else:
                        os.system(f'xdg-open "{file_path}"')
                except Exception:
                    pass
        else:
            messagebox.showerror("Export Error", message)

    def refresh(self):
        """Refresh the page."""
        departments = self.data_manager.departments
        if departments:
            self.dept_dropdown.configure(values=departments)
            if self.dept_var.get() == "Select Department":
                self.dept_var.set(departments[0])
        else:
            self.dept_dropdown.configure(values=["Select Department"])
            self.dept_var.set("Select Department")

        self._update_preview()


# ==================== Main Application ====================

class ExamSeatingApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("University Exam Seating System")
        self.geometry("1200x750")
        self.minsize(1000, 650)

        # Initialize core components
        self.data_manager = DataManager()
        self.algorithm = ZigzagSeatingAlgorithm()
        self.exporter = WordExporter()

        # Configure grid
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create Sidebar
        self.sidebar = Sidebar(self, nav_callback=self._navigate)
        self.sidebar.grid(row=0, column=0, sticky="nsw")

        # Create main content area
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        # Create pages
        self.pages = {}
        self._create_pages()

        # Show dashboard
        self._show_page("dashboard")

    def _create_pages(self):
        """Create all application pages."""
        self.pages["dashboard"] = DashboardPage(
            self.content_frame,
            self.data_manager
        )

        self.pages["departments"] = DepartmentsPage(
            self.content_frame,
            self.data_manager
        )

        self.pages["students"] = StudentsPage(
            self.content_frame,
            self.data_manager
        )

        self.pages["generate"] = GeneratePage(
            self.content_frame,
            self.data_manager,
            self.algorithm
        )

        self.pages["export"] = ExportPage(
            self.content_frame,
            self.data_manager,
            self.algorithm,
            self.exporter
        )

        # Place all pages in the same grid cell
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

    def _show_page(self, key: str):
        """Show a specific page."""
        page = self.pages.get(key)
        if page:
            page.tkraise()
            if hasattr(page, 'refresh'):
                page.refresh()

    def _navigate(self, key: str):
        """Handle navigation."""
        self._show_page(key)


# ==================== Entry Point ====================

def main():
    """Application entry point."""
    app = ExamSeatingApp()
    app.mainloop()


if __name__ == "__main__":
    main()
