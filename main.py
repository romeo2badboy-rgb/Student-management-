#!/usr/bin/env python3
"""
University Exam Seating System - Final Version v2.2.0
A modern desktop application for anti-cheating exam seating management.

Features:
- Multi-department support
- Import students from .docx and .xlsx files
- Manual student entry
- Strict alternating seating algorithm (1,2,1,2 or 1,2,3,1,2,3)
- Professional Word document export with visual seating maps
- Modern dark-themed UI with CustomTkinter
- Clear All functionality for new sessions
- Comprehensive error handling
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Optional, Callable, Dict, Any
import os
import traceback

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
    "primary_light": "#60A5FA",
    "secondary": "#6B7280",
    "success": "#10B981",
    "success_hover": "#059669",
    "warning": "#F59E0B",
    "warning_hover": "#D97706",
    "danger": "#EF4444",
    "danger_hover": "#DC2626",
    "bg_dark": "#1F2937",
    "bg_card": "#374151",
    "bg_hover": "#4B5563",
    "text": "#F9FAFB",
    "text_secondary": "#9CA3AF",
    "border": "#4B5563",
    "accent": "#8B5CF6"
}

# UI Scaling
SCALE = {
    "sidebar_width": 250,
    "card_corner": 12,
    "button_height": 42,
    "entry_height": 42,
    "font_title": 28,
    "font_subtitle": 14,
    "font_normal": 13,
    "font_small": 11,
    "padding_large": 25,
    "padding_medium": 15,
    "padding_small": 8
}


# ==================== Error Handler ====================

def safe_execute(func: Callable, error_title: str = "Error") -> Any:
    """Safely execute a function with error handling."""
    try:
        return func()
    except Exception as e:
        error_msg = str(e)
        print(f"Error: {error_msg}")
        print(traceback.format_exc())
        messagebox.showerror(error_title, f"An error occurred:\n{error_msg}")
        return None


# ==================== Sidebar Navigation ====================

class Sidebar(ctk.CTkFrame):
    """Modern sidebar navigation component."""

    def __init__(self, master, nav_callback: Callable, clear_callback: Callable, **kwargs):
        super().__init__(master, width=SCALE["sidebar_width"], corner_radius=0, **kwargs)

        self.nav_callback = nav_callback
        self.clear_callback = clear_callback
        self.buttons: Dict[str, ctk.CTkButton] = {}
        self.current_page = "dashboard"

        # Prevent sidebar from shrinking
        self.grid_propagate(False)
        self.pack_propagate(False)

        self._create_header()
        self._create_navigation()
        self._create_footer()

    def _create_header(self):
        """Create the header section."""
        # App Header
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=15, pady=(25, 8))

        self.logo_label = ctk.CTkLabel(
            self.header_frame,
            text="🎓",
            font=ctk.CTkFont(size=48)
        )
        self.logo_label.pack()

        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text="Exam Seating",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.title_label.pack(pady=(8, 0))

        self.subtitle_label = ctk.CTkLabel(
            self.header_frame,
            text="Management System",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle_label.pack()

        # Divider
        self.divider = ctk.CTkFrame(self, height=2, fg_color=COLORS["border"])
        self.divider.pack(fill="x", padx=20, pady=20)

    def _create_navigation(self):
        """Create navigation buttons."""
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
                height=48,
                corner_radius=10,
                fg_color="transparent",
                text_color=COLORS["text"],
                hover_color=COLORS["bg_card"],
                anchor="w",
                command=lambda k=key: self._navigate(k)
            )
            btn.pack(fill="x", padx=12, pady=4)
            self.buttons[key] = btn

        # Set initial highlight
        self._highlight("dashboard")

        # Spacer
        self.spacer = ctk.CTkFrame(self, fg_color="transparent")
        self.spacer.pack(fill="both", expand=True)

    def _create_footer(self):
        """Create footer with clear button and theme toggle."""
        # Clear All Button
        self.clear_divider = ctk.CTkFrame(self, height=1, fg_color=COLORS["border"])
        self.clear_divider.pack(fill="x", padx=20, pady=(10, 10))

        self.clear_btn = ctk.CTkButton(
            self,
            text="🗑️  Clear All Data",
            font=ctk.CTkFont(size=13),
            height=40,
            corner_radius=8,
            fg_color=COLORS["danger"],
            hover_color=COLORS["danger_hover"],
            command=self._on_clear_all
        )
        self.clear_btn.pack(fill="x", padx=12, pady=(0, 10))

        # Theme Toggle
        self.theme_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.theme_frame.pack(fill="x", padx=15, pady=(5, 10))

        self.theme_label = ctk.CTkLabel(
            self.theme_frame,
            text="Theme",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_secondary"]
        )
        self.theme_label.pack(anchor="w")

        self.theme_switch = ctk.CTkSwitch(
            self.theme_frame,
            text="Dark Mode",
            font=ctk.CTkFont(size=11),
            command=self._toggle_theme,
            onvalue="dark",
            offvalue="light",
            height=22
        )
        self.theme_switch.pack(anchor="w", pady=3)
        self.theme_switch.select()

        # Version
        self.version_label = ctk.CTkLabel(
            self,
            text="v2.2.0",
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_secondary"]
        )
        self.version_label.pack(pady=(5, 15))

    def _navigate(self, key: str):
        """Handle navigation click."""
        try:
            self.current_page = key
            self._highlight(key)
            self.nav_callback(key)
        except Exception as e:
            print(f"Navigation error: {e}")

    def _highlight(self, key: str):
        """Highlight the active navigation button."""
        for btn_key, btn in self.buttons.items():
            if btn_key == key:
                btn.configure(fg_color=COLORS["primary"])
            else:
                btn.configure(fg_color="transparent")

    def _toggle_theme(self):
        """Toggle between dark and light theme."""
        try:
            mode = self.theme_switch.get()
            ctk.set_appearance_mode(mode)
        except Exception as e:
            print(f"Theme toggle error: {e}")

    def _on_clear_all(self):
        """Handle Clear All button click."""
        self.clear_callback()


# ==================== Dashboard Page ====================

class DashboardPage(ctk.CTkFrame):
    """Main dashboard with overview statistics and quick actions."""

    def __init__(self, master, data_manager: DataManager, nav_callback: Callable = None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.nav_callback = nav_callback

        # Scrollable container
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_frame.pack(fill="both", expand=True)

        self._create_header()
        self._create_university_card()
        self._create_stats_section()
        self._create_quick_actions()
        self._create_status_section()

    def _create_header(self):
        """Create page header."""
        self.title = ctk.CTkLabel(
            self.scroll_frame,
            text="📊 Dashboard",
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self.scroll_frame,
            text="Overview of your exam seating system",
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, SCALE["padding_large"]))

    def _create_university_card(self):
        """Create university name card."""
        self.uni_card = ctk.CTkFrame(self.scroll_frame, corner_radius=SCALE["card_corner"])
        self.uni_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.uni_label = ctk.CTkLabel(
            self.uni_card,
            text="🏫 University Name",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"]
        )
        self.uni_label.pack(anchor="w", padx=20, pady=(15, 5))

        self.uni_entry = ctk.CTkEntry(
            self.uni_card,
            placeholder_text="Enter university name...",
            height=SCALE["entry_height"],
            font=ctk.CTkFont(size=SCALE["font_normal"])
        )
        self.uni_entry.pack(fill="x", padx=20, pady=(0, 15))
        self.uni_entry.insert(0, "University Name")
        self.uni_entry.bind("<KeyRelease>", self._on_uni_change)

    def _create_stats_section(self):
        """Create statistics cards."""
        self.stats_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.stats_frame.pack(fill="x", pady=10)
        self.stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.stat_cards: Dict[str, ctk.CTkFrame] = {}
        stats_config = [
            ("departments", "🏛️", "Departments", "0", COLORS["primary"]),
            ("students", "👥", "Total Students", "0", COLORS["success"]),
            ("stages", "📚", "Active Stages", "0", COLORS["warning"]),
            ("desks", "🪑", "Desks Needed", "0", COLORS["accent"])
        ]

        for i, (key, icon, label, value, color) in enumerate(stats_config):
            card = self._create_stat_card(self.stats_frame, icon, label, value, color)
            card.grid(row=0, column=i, padx=8, pady=5, sticky="nsew")
            self.stat_cards[key] = card

    def _create_stat_card(self, parent, icon: str, label: str, value: str, accent_color: str) -> ctk.CTkFrame:
        """Create a statistics card with accent color."""
        card = ctk.CTkFrame(parent, corner_radius=SCALE["card_corner"], height=140)
        card.pack_propagate(False)

        # Accent bar at top
        accent = ctk.CTkFrame(card, height=4, fg_color=accent_color, corner_radius=2)
        accent.pack(fill="x", padx=10, pady=(10, 0))

        icon_label = ctk.CTkLabel(card, text=icon, font=ctk.CTkFont(size=32))
        icon_label.pack(pady=(12, 5))

        value_label = ctk.CTkLabel(
            card,
            text=value,
            font=ctk.CTkFont(size=28, weight="bold")
        )
        value_label.pack()

        text_label = ctk.CTkLabel(
            card,
            text=label,
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_secondary"]
        )
        text_label.pack(pady=(0, 10))

        card.value_label = value_label
        return card

    def _create_quick_actions(self):
        """Create quick action buttons."""
        self.actions_label = ctk.CTkLabel(
            self.scroll_frame,
            text="⚡ Quick Actions",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.actions_label.pack(anchor="w", pady=(SCALE["padding_large"], SCALE["padding_medium"]))

        self.actions_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.actions_frame.pack(fill="x")

        # Quick action buttons
        actions = [
            ("➕ New Department", COLORS["success"], lambda: self._quick_nav("departments")),
            ("📥 Import Students", COLORS["primary"], lambda: self._quick_nav("students")),
            ("🎲 Generate Seating", COLORS["warning"], lambda: self._quick_nav("generate")),
            ("📄 Export Report", COLORS["accent"], lambda: self._quick_nav("export"))
        ]

        for i, (text, color, command) in enumerate(actions):
            btn = ctk.CTkButton(
                self.actions_frame,
                text=text,
                font=ctk.CTkFont(size=13, weight="bold"),
                height=45,
                width=170,
                fg_color=color,
                hover_color=self._darken_color(color),
                command=command
            )
            btn.grid(row=0, column=i, padx=8, pady=5)

    def _create_status_section(self):
        """Create system status section."""
        self.status_label = ctk.CTkLabel(
            self.scroll_frame,
            text="📋 Quick Start Guide",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.status_label.pack(anchor="w", pady=(SCALE["padding_large"], SCALE["padding_medium"]))

        self.status_card = ctk.CTkFrame(self.scroll_frame, corner_radius=SCALE["card_corner"])
        self.status_card.pack(fill="both", expand=True)

        self.status_text = ctk.CTkTextbox(
            self.status_card,
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            height=200
        )
        self.status_text.pack(fill="both", expand=True, padx=15, pady=15)
        self._update_status_text()
        self.status_text.configure(state="disabled")

    def _darken_color(self, hex_color: str) -> str:
        """Darken a hex color for hover effect."""
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
        """Navigate to a page via quick action."""
        if self.nav_callback:
            self.nav_callback(page)

    def _update_status_text(self):
        """Update the status text."""
        self.status_text.configure(state="normal")
        self.status_text.delete("1.0", "end")

        guide_text = """Welcome to the Exam Seating Management System!

📌 How to Use This Application:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Step 1️⃣  Create Departments
   → Go to "Departments" page
   → Enter department name and click "Create"

Step 2️⃣  Import Students
   → Go to "Students" page
   → Select department and stage
   → Import from Word (.docx) or Excel (.xlsx) file
   → Or add students manually

Step 3️⃣  Generate Seating
   → Go to "Generate Seats" page
   → Select department and click "Generate Seating"
   → Students will be arranged in anti-cheating pattern

Step 4️⃣  Export Report
   → Go to "Export Report" page
   → Configure options and export to Word document

💡 Anti-Cheating Pattern:
   Students are arranged in strict alternating order:
   • 2 stages: 1→2→1→2→1→2...
   • 3 stages: 1→2→3→1→2→3..."""

        self.status_text.insert("1.0", guide_text)
        self.status_text.configure(state="disabled")

    def _on_uni_change(self, event):
        """Update university name."""
        try:
            self.data_manager.university_name = self.uni_entry.get()
        except Exception as e:
            print(f"Error updating university name: {e}")

    def refresh(self):
        """Refresh dashboard statistics."""
        try:
            # Update university name
            current_uni = self.data_manager.university_name
            if self.uni_entry.get() != current_uni:
                self.uni_entry.delete(0, "end")
                self.uni_entry.insert(0, current_uni)

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

            # Calculate desks needed (2 students per desk, round up)
            desks_needed = (total_students + 1) // 2
            self.stat_cards["desks"].value_label.configure(text=str(desks_needed))
        except Exception as e:
            print(f"Error refreshing dashboard: {e}")


# ==================== Departments Page ====================

class DepartmentsPage(ctk.CTkFrame):
    """Department management page."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.dept_widgets = []

        self._create_header()
        self._create_add_section()
        self._create_list_section()

    def _create_header(self):
        """Create page header."""
        self.title = ctk.CTkLabel(
            self,
            text="🏛️ Departments",
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Manage university departments",
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, SCALE["padding_large"]))

    def _create_add_section(self):
        """Create add department section."""
        self.add_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.add_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.add_label = ctk.CTkLabel(
            self.add_card,
            text="➕ Create New Department",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.add_label.pack(anchor="w", padx=20, pady=(15, 10))

        self.add_frame = ctk.CTkFrame(self.add_card, fg_color="transparent")
        self.add_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.dept_entry = ctk.CTkEntry(
            self.add_frame,
            placeholder_text="Enter department name (e.g., Computer Science)",
            height=SCALE["entry_height"],
            font=ctk.CTkFont(size=SCALE["font_normal"])
        )
        self.dept_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.dept_entry.bind("<Return>", lambda e: self._create_department())

        self.add_btn = ctk.CTkButton(
            self.add_frame,
            text="➕ Create",
            font=ctk.CTkFont(size=SCALE["font_normal"], weight="bold"),
            height=SCALE["button_height"],
            width=130,
            fg_color=COLORS["success"],
            hover_color=COLORS["success_hover"],
            command=self._create_department
        )
        self.add_btn.pack(side="right")

    def _create_list_section(self):
        """Create departments list section."""
        self.list_label = ctk.CTkLabel(
            self,
            text="📋 Existing Departments",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.list_label.pack(anchor="w", pady=(10, SCALE["padding_medium"]))

        self.list_frame = ctk.CTkScrollableFrame(self, corner_radius=SCALE["card_corner"])
        self.list_frame.pack(fill="both", expand=True)

    def _create_department(self):
        """Create a new department."""
        try:
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
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create department: {e}")

    def _delete_department(self, name: str):
        """Delete a department."""
        try:
            if messagebox.askyesno(
                "Confirm Delete",
                f"Delete department '{name}'?\n\nThis will remove all student data for this department."
            ):
                success, message = self.data_manager.delete_department(name)
                if success:
                    self.refresh()
                    messagebox.showinfo("Success", message)
                else:
                    messagebox.showerror("Error", message)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete department: {e}")

    def _create_dept_card(self, parent, name: str, stats: dict) -> ctk.CTkFrame:
        """Create a department card widget."""
        card = ctk.CTkFrame(parent, corner_radius=10)

        # Main content frame
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="x", padx=15, pady=12)

        # Left side - info
        info_frame = ctk.CTkFrame(content, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True)

        name_label = ctk.CTkLabel(
            info_frame,
            text=f"🏛️ {name}",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        name_label.pack(anchor="w")

        # Stats line
        stages = stats.get("stages", {})
        stage_counts = [f"{s}: {c}" for s, c in stages.items() if c > 0]
        total = stats.get("total_students", 0)

        stats_text = f"Total: {total} students"
        if stage_counts:
            stats_text += f"  •  " + " | ".join(stage_counts)

        stats_label = ctk.CTkLabel(
            info_frame,
            text=stats_text,
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_secondary"],
            anchor="w"
        )
        stats_label.pack(anchor="w", pady=(3, 0))

        # Right side - delete button
        del_btn = ctk.CTkButton(
            content,
            text="🗑️ Delete",
            width=90,
            height=34,
            font=ctk.CTkFont(size=11),
            fg_color=COLORS["danger"],
            hover_color=COLORS["danger_hover"],
            command=lambda: self._delete_department(name)
        )
        del_btn.pack(side="right")

        return card

    def refresh(self):
        """Refresh the departments list."""
        try:
            # Clear existing widgets
            for widget in self.dept_widgets:
                widget.destroy()
            self.dept_widgets.clear()

            # Create new cards
            departments = self.data_manager.departments
            if not departments:
                empty_label = ctk.CTkLabel(
                    self.list_frame,
                    text="No departments created yet.\n\nCreate your first department above to get started.",
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


# ==================== Students Page ====================

class StudentsPage(ctk.CTkFrame):
    """Student management and import page."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.stage_textboxes: Dict[str, tuple] = {}

        self._create_header()
        self._create_import_section()
        self._create_manual_entry_section()
        self._create_list_section()

    def _create_header(self):
        """Create page header."""
        self.title = ctk.CTkLabel(
            self,
            text="👥 Students",
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Import and manage student lists",
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, SCALE["padding_large"]))

    def _create_import_section(self):
        """Create import section."""
        self.import_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.import_card.pack(fill="x", pady=(0, SCALE["padding_medium"]))

        self.import_title = ctk.CTkLabel(
            self.import_card,
            text="📥 Import Students from File",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.import_title.pack(anchor="w", padx=20, pady=(15, 10))

        # Selection Row
        self.select_frame = ctk.CTkFrame(self.import_card, fg_color="transparent")
        self.select_frame.pack(fill="x", padx=20, pady=5)

        self.dept_label = ctk.CTkLabel(
            self.select_frame,
            text="Department:",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.dept_label.pack(side="left", padx=(0, 8))

        self.dept_var = ctk.StringVar(value="Select Department")
        self.dept_dropdown = ctk.CTkOptionMenu(
            self.select_frame,
            variable=self.dept_var,
            values=["Select Department"],
            width=200,
            height=36,
            font=ctk.CTkFont(size=12),
            command=self._on_dept_change
        )
        self.dept_dropdown.pack(side="left", padx=(0, 20))

        self.stage_label = ctk.CTkLabel(
            self.select_frame,
            text="Stage:",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.stage_label.pack(side="left", padx=(0, 8))

        self.stage_var = ctk.StringVar(value="1st Stage")
        self.stage_dropdown = ctk.CTkOptionMenu(
            self.select_frame,
            variable=self.stage_var,
            values=["1st Stage", "2nd Stage", "3rd Stage"],
            width=140,
            height=36,
            font=ctk.CTkFont(size=12)
        )
        self.stage_dropdown.pack(side="left")

        # Import Buttons Row
        self.btn_frame = ctk.CTkFrame(self.import_card, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=20, pady=(10, 15))

        self.import_docx_btn = ctk.CTkButton(
            self.btn_frame,
            text="📄 Import Word (.docx)",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=40,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            command=lambda: self._import_file("docx")
        )
        self.import_docx_btn.pack(side="left", padx=(0, 10))

        self.import_xlsx_btn = ctk.CTkButton(
            self.btn_frame,
            text="📊 Import Excel (.xlsx)",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=40,
            fg_color=COLORS["success"],
            hover_color=COLORS["success_hover"],
            command=lambda: self._import_file("xlsx")
        )
        self.import_xlsx_btn.pack(side="left", padx=(0, 10))

        self.clear_stage_btn = ctk.CTkButton(
            self.btn_frame,
            text="🗑️ Clear Stage",
            font=ctk.CTkFont(size=12),
            height=40,
            fg_color=COLORS["danger"],
            hover_color=COLORS["danger_hover"],
            command=self._clear_stage
        )
        self.clear_stage_btn.pack(side="left")

    def _create_manual_entry_section(self):
        """Create manual student entry section."""
        self.manual_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.manual_card.pack(fill="x", pady=(0, SCALE["padding_medium"]))

        self.manual_title = ctk.CTkLabel(
            self.manual_card,
            text="✏️ Add Student Manually",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.manual_title.pack(anchor="w", padx=20, pady=(15, 10))

        self.manual_frame = ctk.CTkFrame(self.manual_card, fg_color="transparent")
        self.manual_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.student_entry = ctk.CTkEntry(
            self.manual_frame,
            placeholder_text="Enter student name...",
            height=SCALE["entry_height"],
            font=ctk.CTkFont(size=SCALE["font_normal"])
        )
        self.student_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.student_entry.bind("<Return>", lambda e: self._add_student_manually())

        self.add_student_btn = ctk.CTkButton(
            self.manual_frame,
            text="➕ Add Student",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=SCALE["button_height"],
            width=130,
            fg_color=COLORS["accent"],
            hover_color="#7C3AED",
            command=self._add_student_manually
        )
        self.add_student_btn.pack(side="right")

    def _create_list_section(self):
        """Create student list section."""
        self.list_label = ctk.CTkLabel(
            self,
            text="📋 Student Lists by Stage",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.list_label.pack(anchor="w", pady=(10, SCALE["padding_medium"]))

        # Tabview for stages
        self.tabview = ctk.CTkTabview(self, corner_radius=SCALE["card_corner"])
        self.tabview.pack(fill="both", expand=True)

        self.stage_tabs = {}

        for stage in ["1st Stage", "2nd Stage", "3rd Stage"]:
            tab = self.tabview.add(stage)
            self.stage_tabs[stage] = tab

            # Count label
            count_label = ctk.CTkLabel(
                tab,
                text="Students: 0",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=COLORS["primary_light"]
            )
            count_label.pack(anchor="e", padx=10, pady=(5, 0))

            textbox = ctk.CTkTextbox(tab, font=ctk.CTkFont(size=12))
            textbox.pack(fill="both", expand=True, padx=10, pady=(5, 10))
            self.stage_textboxes[stage] = (textbox, count_label)

    def _on_dept_change(self, value):
        """Handle department change."""
        self.refresh()

    def _add_student_manually(self):
        """Add a student manually."""
        try:
            dept = self.dept_var.get()
            if dept == "Select Department":
                messagebox.showwarning("Warning", "Please select a department first.")
                return

            name = self.student_entry.get().strip()
            if not name:
                messagebox.showwarning("Warning", "Please enter a student name.")
                return

            stage = self.stage_var.get()
            success = self.data_manager.add_student(dept, stage, name)

            if success:
                self.student_entry.delete(0, "end")
                self.refresh()
                # Don't show message for each student to avoid annoyance
            else:
                messagebox.showinfo("Info", f"Student '{name}' already exists in {stage}.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add student: {e}")

    def _import_file(self, file_type: str):
        """Import students from file."""
        try:
            dept = self.dept_var.get()
            if dept == "Select Department":
                messagebox.showwarning("Warning", "Please select a department first.")
                return

            stage = self.stage_var.get()

            if file_type == "docx":
                filetypes = [("Word Documents", "*.docx"), ("All Files", "*.*")]
            else:
                filetypes = [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]

            file_path = filedialog.askopenfilename(
                title=f"Select {file_type.upper()} File",
                filetypes=filetypes
            )

            if not file_path:
                return

            success, message, count = self.data_manager.import_from_file(file_path, dept, stage)

            if success:
                self.refresh()
                messagebox.showinfo("Import Successful", message)
            else:
                messagebox.showerror("Import Error", message)
        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {e}")

    def _clear_stage(self):
        """Clear students from selected stage."""
        try:
            dept = self.dept_var.get()
            if dept == "Select Department":
                messagebox.showwarning("Warning", "Please select a department.")
                return

            stage = self.stage_var.get()
            count = self.data_manager.get_student_count(dept, stage)

            if count == 0:
                messagebox.showinfo("Info", f"No students in {stage} to clear.")
                return

            if messagebox.askyesno("Confirm Clear", f"Clear all {count} students from {stage}?"):
                self.data_manager.clear_stage(dept, stage)
                self.refresh()
                messagebox.showinfo("Success", f"Cleared {count} students from {stage}.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear stage: {e}")

    def refresh(self):
        """Refresh the students view."""
        try:
            # Update department dropdown
            departments = self.data_manager.departments
            if departments:
                self.dept_dropdown.configure(values=departments)
                if self.dept_var.get() == "Select Department" or self.dept_var.get() not in departments:
                    self.dept_var.set(departments[0])
            else:
                self.dept_dropdown.configure(values=["Select Department"])
                self.dept_var.set("Select Department")

            # Update student lists
            dept = self.dept_var.get()
            if dept != "Select Department":
                for stage, (textbox, count_label) in self.stage_textboxes.items():
                    students = self.data_manager.get_students(dept, stage)

                    textbox.configure(state="normal")
                    textbox.delete("1.0", "end")

                    if students:
                        for i, student in enumerate(students, 1):
                            textbox.insert("end", f"{i:3}. {student}\n")
                        count_label.configure(text=f"Students: {len(students)}")
                    else:
                        textbox.insert("1.0", "No students imported for this stage.\n\nUse the import buttons or manual entry above to add students.")
                        count_label.configure(text="Students: 0")

                    textbox.configure(state="disabled")
            else:
                # Clear all textboxes if no department selected
                for stage, (textbox, count_label) in self.stage_textboxes.items():
                    textbox.configure(state="normal")
                    textbox.delete("1.0", "end")
                    textbox.insert("1.0", "Please select a department first.")
                    textbox.configure(state="disabled")
                    count_label.configure(text="Students: 0")
        except Exception as e:
            print(f"Error refreshing students: {e}")


# ==================== Generate Seating Page ====================

class GeneratePage(ctk.CTkFrame):
    """Seating generation page."""

    def __init__(self, master, data_manager: DataManager, algorithm: ZigzagSeatingAlgorithm, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.data_manager = data_manager
        self.algorithm = algorithm

        self._create_header()
        self._create_info_section()
        self._create_controls()
        self._create_results_section()

    def _create_header(self):
        """Create page header."""
        self.title = ctk.CTkLabel(
            self,
            text="🎲 Generate Seating",
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Create anti-cheating seating arrangements with strict alternation",
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, SCALE["padding_large"]))

    def _create_info_section(self):
        """Create algorithm info section."""
        self.info_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.info_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.info_title = ctk.CTkLabel(
            self.info_card,
            text="🛡️ Anti-Cheating Algorithm",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.info_title.pack(anchor="w", padx=20, pady=(15, 8))

        info_text = (
            "• STRICT Alternation: 2 stages → (1,2,1,2...) | 3 stages → (1,2,3,1,2,3...)\n"
            "• Students within each stage are RANDOMIZED for fairness\n"
            "• Desks arranged in 3 columns (Right, Middle, Left)\n"
            "• Empty seats marked as '--- Empty ---' if odd count"
        )
        self.info_label = ctk.CTkLabel(
            self.info_card,
            text=info_text,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"],
            justify="left"
        )
        self.info_label.pack(anchor="w", padx=20, pady=(0, 15))

    def _create_controls(self):
        """Create generation controls."""
        self.control_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.control_frame.pack(fill="x", pady=(0, SCALE["padding_medium"]))

        self.dept_label = ctk.CTkLabel(
            self.control_frame,
            text="Select Department:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.dept_label.pack(side="left", padx=(0, 10))

        self.dept_var = ctk.StringVar(value="Select Department")
        self.dept_dropdown = ctk.CTkOptionMenu(
            self.control_frame,
            variable=self.dept_var,
            values=["Select Department"],
            width=220,
            height=42,
            font=ctk.CTkFont(size=12)
        )
        self.dept_dropdown.pack(side="left", padx=(0, 15))

        self.generate_btn = ctk.CTkButton(
            self.control_frame,
            text="🎲 Generate Seating",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=42,
            width=180,
            fg_color=COLORS["warning"],
            hover_color=COLORS["warning_hover"],
            command=self._generate
        )
        self.generate_btn.pack(side="left")

    def _create_results_section(self):
        """Create results display section."""
        self.results_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.results_card.pack(fill="both", expand=True)

        self.results_title = ctk.CTkLabel(
            self.results_card,
            text="📊 Generated Seating Arrangement",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.results_title.pack(anchor="w", padx=20, pady=(15, 10))

        self.results_text = ctk.CTkTextbox(
            self.results_card,
            font=ctk.CTkFont(size=11, family="Courier")
        )
        self.results_text.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self._show_initial_message()

        # Stats bar
        self.stats_label = ctk.CTkLabel(
            self.results_card,
            text="",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COLORS["success"]
        )
        self.stats_label.pack(anchor="w", padx=20, pady=(0, 15))

    def _show_initial_message(self):
        """Show initial message in results area."""
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")
        self.results_text.insert("1.0", "Select a department and click 'Generate Seating' to begin...\n\n")
        self.results_text.insert("end", "The algorithm will:\n")
        self.results_text.insert("end", "  ✓ Alternate students strictly between stages\n")
        self.results_text.insert("end", "  ✓ Randomize students within each stage\n")
        self.results_text.insert("end", "  ✓ Handle odd student counts with empty seat markers\n")
        self.results_text.insert("end", "  ✓ Maximize cross-stage pairs to prevent cheating\n")
        self.results_text.configure(state="disabled")

    def _generate(self):
        """Generate seating arrangement."""
        try:
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

            self._display_results(result)

            # Show warnings if any
            if result.warnings:
                messagebox.showwarning("Notice", "\n".join(result.warnings))
            else:
                messagebox.showinfo("Success", message)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate seating: {e}")
            print(traceback.format_exc())

    def _display_results(self, result: SeatingResult):
        """Display the generated seating."""
        try:
            self.results_text.configure(state="normal")
            self.results_text.delete("1.0", "end")

            # Header with pattern info
            pattern = " → ".join(result.stage_order) if result.stage_order else "N/A"
            self.results_text.insert("end", f"Seating Pattern: {pattern} (repeating)\n")
            self.results_text.insert("end", "=" * 110 + "\n\n")

            # Column headers
            header = f"{'Desk':<6} {'Col':<8} {'Row':<5} {'Seat A':<8} {'Student A':<30} {'Seat B':<8} {'Student B':<30}\n"
            self.results_text.insert("end", header)
            self.results_text.insert("end", "-" * 110 + "\n")

            # Data rows
            for desk in result.desks:
                # Student A
                if desk.student_a:
                    if desk.student_a.is_empty:
                        student_a = "--- Empty ---"
                    else:
                        name_a = desk.student_a.name[:22] + "..." if len(desk.student_a.name) > 22 else desk.student_a.name
                        student_a = f"{name_a} ({desk.student_a.stage[:3]})"
                else:
                    student_a = "-"

                # Student B
                if desk.student_b:
                    if desk.student_b.is_empty:
                        student_b = "--- Empty ---"
                    else:
                        name_b = desk.student_b.name[:22] + "..." if len(desk.student_b.name) > 22 else desk.student_b.name
                        student_b = f"{name_b} ({desk.student_b.stage[:3]})"
                else:
                    student_b = "-"

                row = f"{desk.number:<6} {desk.column.value:<8} {desk.row:<5} {desk.seat_a_number:<8} {student_a:<30} {desk.seat_b_number:<8} {student_b:<30}\n"
                self.results_text.insert("end", row)

            self.results_text.configure(state="disabled")

            # Update stats
            stats = result.stats
            empty_count = stats.get('empty_seats', 0)
            empty_text = f" | Empty Seats: {empty_count}" if empty_count > 0 else ""

            self.stats_label.configure(
                text=f"✅ {stats.get('total_desks', 0)} desks | "
                f"{stats.get('total_students', 0)} students | "
                f"Cross-stage pairs: {stats.get('cross_stage_pairs', 0)}{empty_text}"
            )
        except Exception as e:
            print(f"Error displaying results: {e}")

    def refresh(self):
        """Refresh the page."""
        try:
            departments = self.data_manager.departments
            if departments:
                self.dept_dropdown.configure(values=departments)
                if self.dept_var.get() == "Select Department" or self.dept_var.get() not in departments:
                    self.dept_var.set(departments[0])
            else:
                self.dept_dropdown.configure(values=["Select Department"])
                self.dept_var.set("Select Department")
        except Exception as e:
            print(f"Error refreshing generate page: {e}")


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

        self._create_header()
        self._create_options_section()
        self._create_export_button()
        self._create_preview_section()

    def _create_header(self):
        """Create page header."""
        self.title = ctk.CTkLabel(
            self,
            text="📄 Export Report",
            font=ctk.CTkFont(size=SCALE["font_title"], weight="bold")
        )
        self.title.pack(anchor="w", pady=(0, 5))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Generate professional Word documents with visual seating maps",
            font=ctk.CTkFont(size=SCALE["font_subtitle"]),
            text_color=COLORS["text_secondary"]
        )
        self.subtitle.pack(anchor="w", pady=(0, SCALE["padding_large"]))

    def _create_options_section(self):
        """Create export options section."""
        self.options_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.options_card.pack(fill="x", pady=(0, SCALE["padding_large"]))

        self.options_title = ctk.CTkLabel(
            self.options_card,
            text="⚙️ Export Options",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.options_title.pack(anchor="w", padx=20, pady=(15, 15))

        # Options grid
        options_frame = ctk.CTkFrame(self.options_card, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=(0, 15))

        # Row 1: Department
        ctk.CTkLabel(
            options_frame,
            text="Department:",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=100
        ).grid(row=0, column=0, sticky="w", pady=8)

        self.dept_var = ctk.StringVar(value="Select Department")
        self.dept_dropdown = ctk.CTkOptionMenu(
            options_frame,
            variable=self.dept_var,
            values=["Select Department"],
            width=250,
            height=36
        )
        self.dept_dropdown.grid(row=0, column=1, sticky="w", pady=8, padx=(10, 0))

        # Row 2: Exam Title
        ctk.CTkLabel(
            options_frame,
            text="Exam Title:",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=100
        ).grid(row=1, column=0, sticky="w", pady=8)

        self.title_entry = ctk.CTkEntry(
            options_frame,
            placeholder_text="e.g., Final Exam Seating",
            width=400,
            height=36
        )
        self.title_entry.grid(row=1, column=1, sticky="w", pady=8, padx=(10, 0))
        self.title_entry.insert(0, "Exam Seating Arrangement")

        # Row 3: Date
        ctk.CTkLabel(
            options_frame,
            text="Exam Date:",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=100
        ).grid(row=2, column=0, sticky="w", pady=8)

        self.date_entry = ctk.CTkEntry(
            options_frame,
            placeholder_text="Leave empty for today's date",
            width=250,
            height=36
        )
        self.date_entry.grid(row=2, column=1, sticky="w", pady=8, padx=(10, 0))

        # Row 4: Options
        ctk.CTkLabel(
            options_frame,
            text="Options:",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=100
        ).grid(row=3, column=0, sticky="w", pady=8)

        self.map_var = ctk.BooleanVar(value=True)
        self.map_check = ctk.CTkCheckBox(
            options_frame,
            text="Include Visual Seating Map",
            variable=self.map_var,
            font=ctk.CTkFont(size=12)
        )
        self.map_check.grid(row=3, column=1, sticky="w", pady=8, padx=(10, 0))

    def _create_export_button(self):
        """Create export button."""
        self.export_btn = ctk.CTkButton(
            self,
            text="📄 Export to Word Document",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=50,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            command=self._export
        )
        self.export_btn.pack(pady=SCALE["padding_large"])

    def _create_preview_section(self):
        """Create preview section."""
        self.preview_card = ctk.CTkFrame(self, corner_radius=SCALE["card_corner"])
        self.preview_card.pack(fill="both", expand=True)

        self.preview_title = ctk.CTkLabel(
            self.preview_card,
            text="👁️ Export Preview",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.preview_title.pack(anchor="w", padx=20, pady=(15, 10))

        self.preview_text = ctk.CTkTextbox(self.preview_card, font=ctk.CTkFont(size=12))
        self.preview_text.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        self._update_preview()
        self.preview_text.configure(state="disabled")

    def _update_preview(self):
        """Update the preview text."""
        try:
            self.preview_text.configure(state="normal")
            self.preview_text.delete("1.0", "end")

            result = self.algorithm.result
            if not result or not result.desks:
                self.preview_text.insert("1.0", "⚠️ No seating arrangement generated yet.\n\n")
                self.preview_text.insert("end", "Please go to 'Generate Seats' page first to create\n")
                self.preview_text.insert("end", "a seating arrangement before exporting.\n\n")
                self.preview_text.insert("end", "The exported document will include:\n")
                self.preview_text.insert("end", "  • University and department header\n")
                self.preview_text.insert("end", "  • Detailed seating table with seat numbers\n")
                self.preview_text.insert("end", "  • Visual seating map (desk layout)\n")
                self.preview_text.insert("end", "  • Statistics and signature section\n")
            else:
                self.preview_text.insert("1.0", "✅ Seating arrangement ready for export!\n\n")
                self.preview_text.insert("end", f"📊 Summary:\n")
                self.preview_text.insert("end", f"   • Total Desks: {result.total_desks}\n")
                self.preview_text.insert("end", f"   • Total Students: {result.total_students}\n")
                self.preview_text.insert("end", f"   • Cross-Stage Pairs: {result.cross_stage_count}\n")
                self.preview_text.insert("end", f"   • Empty Seats: {result.empty_seats_count}\n")

                if result.stage_order:
                    self.preview_text.insert("end", f"   • Pattern: {' → '.join(result.stage_order)}\n")

                if result.warnings:
                    self.preview_text.insert("end", f"\n⚠️ Warnings:\n")
                    for warning in result.warnings:
                        self.preview_text.insert("end", f"   • {warning}\n")

            self.preview_text.configure(state="disabled")
        except Exception as e:
            print(f"Error updating preview: {e}")

    def _export(self):
        """Export the report."""
        try:
            result = self.algorithm.result
            if not result or not result.desks:
                messagebox.showerror("Error", "No seating arrangement to export.\nPlease generate seating first.")
                return

            dept_name = self.dept_var.get()
            if dept_name == "Select Department":
                messagebox.showwarning("Warning", "Please select a department.")
                return

            # Get options
            exam_title = self.title_entry.get().strip() or "Exam Seating Arrangement"
            exam_date = self.date_entry.get().strip() or None
            include_map = self.map_var.get()

            # Get save path
            default_name = f"{dept_name.replace(' ', '_')}_seating.docx"
            file_path = filedialog.asksaveasfilename(
                title="Save Word Document",
                defaultextension=".docx",
                filetypes=[("Word Documents", "*.docx")],
                initialfile=default_name
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
                messagebox.showinfo("Export Successful", message)

                if messagebox.askyesno("Open File", "Would you like to open the exported file?"):
                    try:
                        if os.name == 'nt':
                            os.startfile(file_path)
                        else:
                            os.system(f'xdg-open "{file_path}" &')
                    except Exception:
                        pass
            else:
                messagebox.showerror("Export Error", message)

        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            print(traceback.format_exc())

    def refresh(self):
        """Refresh the page."""
        try:
            departments = self.data_manager.departments
            if departments:
                self.dept_dropdown.configure(values=departments)
                if self.dept_var.get() == "Select Department" or self.dept_var.get() not in departments:
                    self.dept_var.set(departments[0])
            else:
                self.dept_dropdown.configure(values=["Select Department"])
                self.dept_var.set("Select Department")

            self._update_preview()
        except Exception as e:
            print(f"Error refreshing export page: {e}")


# ==================== Main Application ====================

class ExamSeatingApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("University Exam Seating System v2.2.0")
        self.geometry("1300x850")
        self.minsize(1100, 700)

        # Center window on screen
        self._center_window()

        # Initialize core components
        self.data_manager = DataManager()
        self.algorithm = ZigzagSeatingAlgorithm()
        self.exporter = WordExporter()

        # Configure grid
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create UI
        self._create_sidebar()
        self._create_content_area()
        self._create_pages()

        # Show dashboard
        self._show_page("dashboard")

    def _center_window(self):
        """Center the window on screen."""
        try:
            self.update_idletasks()
            width = 1300
            height = 850
            x = (self.winfo_screenwidth() // 2) - (width // 2)
            y = (self.winfo_screenheight() // 2) - (height // 2)
            self.geometry(f'{width}x{height}+{x}+{y}')
        except Exception:
            pass

    def _create_sidebar(self):
        """Create sidebar navigation."""
        self.sidebar = Sidebar(
            self,
            nav_callback=self._navigate,
            clear_callback=self._clear_all_data
        )
        self.sidebar.grid(row=0, column=0, sticky="nsw")

    def _create_content_area(self):
        """Create main content area."""
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

    def _create_pages(self):
        """Create all application pages."""
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

        # Place all pages in the same grid cell
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

    def _show_page(self, key: str):
        """Show a specific page."""
        try:
            page = self.pages.get(key)
            if page:
                page.tkraise()
                if hasattr(page, 'refresh'):
                    page.refresh()
        except Exception as e:
            print(f"Error showing page {key}: {e}")

    def _navigate(self, key: str):
        """Handle navigation."""
        self._show_page(key)
        self.sidebar._highlight(key)

    def _clear_all_data(self):
        """Clear all data and reset for new session."""
        try:
            if messagebox.askyesno(
                "Clear All Data",
                "Are you sure you want to clear ALL data?\n\n"
                "This will delete:\n"
                "• All departments\n"
                "• All student lists\n"
                "• Generated seating arrangements\n\n"
                "This action cannot be undone!"
            ):
                # Reset data manager
                self.data_manager = DataManager()

                # Clear algorithm result
                self.algorithm.clear()

                # Recreate pages with fresh data
                for page in self.pages.values():
                    page.destroy()
                self.pages.clear()

                self._create_pages()
                self._show_page("dashboard")

                messagebox.showinfo("Data Cleared", "All data has been cleared.\nReady for a new session!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear data: {e}")


# ==================== Entry Point ====================

def main():
    """Application entry point."""
    try:
        app = ExamSeatingApp()
        app.mainloop()
    except Exception as e:
        print(f"Application error: {e}")
        print(traceback.format_exc())
        messagebox.showerror("Fatal Error", f"Application failed to start:\n{e}")


if __name__ == "__main__":
    main()
