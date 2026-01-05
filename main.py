#!/usr/bin/env python3
"""
Exam Seating Management Application
A modern desktop application for managing exam seating arrangements.

Features:
- Import students from Word documents
- Smart cross-stage seating algorithm
- Export to formatted Word documents
- Dark/Light mode theme support
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from typing import Optional

from data_manager import DataManager
from seating_algorithm import SeatingAlgorithm
from word_exporter import WordExporter


class SidebarFrame(ctk.CTkFrame):
    """Sidebar navigation frame with theme toggle."""

    def __init__(self, master, nav_callback, **kwargs):
        super().__init__(master, **kwargs)

        self.nav_callback = nav_callback

        # App Logo/Title
        self.logo_label = ctk.CTkLabel(
            self,
            text="📋 Exam Seating",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.logo_label.pack(pady=(20, 10))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Management System",
            font=ctk.CTkFont(size=12)
        )
        self.subtitle.pack(pady=(0, 30))

        # Navigation Buttons
        self.nav_buttons = {}

        nav_items = [
            ("🏠 Home", "home"),
            ("📥 Import Students", "import"),
            ("📊 View Data", "view"),
            ("🪑 Generate Seats", "generate"),
            ("📄 Export", "export"),
        ]

        for text, key in nav_items:
            btn = ctk.CTkButton(
                self,
                text=text,
                font=ctk.CTkFont(size=14),
                height=40,
                corner_radius=8,
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray70", "gray30"),
                anchor="w",
                command=lambda k=key: self._on_nav_click(k)
            )
            btn.pack(fill="x", padx=10, pady=5)
            self.nav_buttons[key] = btn

        # Highlight home by default
        self._highlight_button("home")

        # Spacer
        self.spacer = ctk.CTkLabel(self, text="")
        self.spacer.pack(expand=True)

        # Theme Toggle
        self.theme_label = ctk.CTkLabel(
            self,
            text="Appearance Mode",
            font=ctk.CTkFont(size=12)
        )
        self.theme_label.pack(pady=(10, 5))

        self.theme_menu = ctk.CTkOptionMenu(
            self,
            values=["System", "Light", "Dark"],
            command=self._change_theme,
            width=140
        )
        self.theme_menu.pack(pady=(0, 20))

    def _on_nav_click(self, key: str):
        """Handle navigation button click."""
        self._highlight_button(key)
        self.nav_callback(key)

    def _highlight_button(self, key: str):
        """Highlight the selected navigation button."""
        for btn_key, btn in self.nav_buttons.items():
            if btn_key == key:
                btn.configure(fg_color=("gray75", "gray25"))
            else:
                btn.configure(fg_color="transparent")

    def _change_theme(self, mode: str):
        """Change the application theme."""
        ctk.set_appearance_mode(mode.lower())


class HomeFrame(ctk.CTkFrame):
    """Home page with overview and department settings."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, **kwargs)

        self.data_manager = data_manager

        # Title
        self.title = ctk.CTkLabel(
            self,
            text="Welcome to Exam Seating Manager",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title.pack(pady=(40, 10))

        self.description = ctk.CTkLabel(
            self,
            text="Organize exam seating with smart cross-stage pairing",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        self.description.pack(pady=(0, 40))

        # Department Name Entry
        self.dept_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.dept_frame.pack(pady=20, padx=40, fill="x")

        self.dept_label = ctk.CTkLabel(
            self.dept_frame,
            text="Department Name:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.dept_label.pack(anchor="w")

        self.dept_entry = ctk.CTkEntry(
            self.dept_frame,
            placeholder_text="e.g., Cyber Security",
            height=40,
            font=ctk.CTkFont(size=14)
        )
        self.dept_entry.pack(fill="x", pady=(5, 0))
        self.dept_entry.bind("<KeyRelease>", self._on_dept_change)

        # Quick Stats
        self.stats_frame = ctk.CTkFrame(self)
        self.stats_frame.pack(pady=30, padx=40, fill="x")

        self.stats_title = ctk.CTkLabel(
            self.stats_frame,
            text="📊 Current Data Summary",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.stats_title.pack(pady=(15, 10))

        self.stats_label = ctk.CTkLabel(
            self.stats_frame,
            text="No students imported yet.\nUse the Import section to add students.",
            font=ctk.CTkFont(size=13),
            justify="center"
        )
        self.stats_label.pack(pady=(0, 15))

        # Instructions
        self.instructions_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.instructions_frame.pack(pady=20, padx=40, fill="x")

        instructions = [
            "1️⃣  Enter your department name above",
            "2️⃣  Import students from .docx files for each stage",
            "3️⃣  Generate smart seating arrangements",
            "4️⃣  Export the results to a Word document"
        ]

        self.inst_title = ctk.CTkLabel(
            self.instructions_frame,
            text="Quick Start Guide:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.inst_title.pack(anchor="w", pady=(0, 10))

        for inst in instructions:
            lbl = ctk.CTkLabel(
                self.instructions_frame,
                text=inst,
                font=ctk.CTkFont(size=13),
                anchor="w"
            )
            lbl.pack(anchor="w", pady=2)

    def _on_dept_change(self, event):
        """Update department name in data manager."""
        self.data_manager.department_name = self.dept_entry.get()

    def refresh(self):
        """Refresh the stats display."""
        # Update department entry if it was set elsewhere
        current_dept = self.data_manager.department_name
        if self.dept_entry.get() != current_dept:
            self.dept_entry.delete(0, "end")
            if current_dept:
                self.dept_entry.insert(0, current_dept)

        # Update stats
        total = self.data_manager.get_total_students()
        if total > 0:
            summary = self.data_manager.get_summary()
            self.stats_label.configure(text=summary)
        else:
            self.stats_label.configure(
                text="No students imported yet.\nUse the Import section to add students."
            )


class ImportFrame(ctk.CTkFrame):
    """Frame for importing students from Word documents."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, **kwargs)

        self.data_manager = data_manager

        # Title
        self.title = ctk.CTkLabel(
            self,
            text="📥 Import Students",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title.pack(pady=(40, 10))

        self.description = ctk.CTkLabel(
            self,
            text="Import student names from Word documents (.docx)",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        self.description.pack(pady=(0, 30))

        # Stage Selection
        self.stage_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.stage_frame.pack(pady=10, padx=40, fill="x")

        self.stage_label = ctk.CTkLabel(
            self.stage_frame,
            text="Select Academic Year:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.stage_label.pack(anchor="w")

        self.stage_var = ctk.StringVar(value="1st Stage")
        self.stage_dropdown = ctk.CTkOptionMenu(
            self.stage_frame,
            variable=self.stage_var,
            values=self.data_manager.stages,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14)
        )
        self.stage_dropdown.pack(anchor="w", pady=(5, 0))

        # Import Button
        self.import_btn = ctk.CTkButton(
            self,
            text="📂 Select & Import .docx File",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            width=250,
            command=self._import_file
        )
        self.import_btn.pack(pady=30)

        # Status/Result
        self.status_frame = ctk.CTkFrame(self)
        self.status_frame.pack(pady=20, padx=40, fill="both", expand=True)

        self.status_title = ctk.CTkLabel(
            self.status_frame,
            text="Import Status",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.status_title.pack(pady=(15, 10))

        self.status_text = ctk.CTkTextbox(
            self.status_frame,
            height=200,
            font=ctk.CTkFont(size=12)
        )
        self.status_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.status_text.insert("1.0", "Ready to import students...\n")
        self.status_text.configure(state="disabled")

        # Clear Stage Button
        self.clear_btn = ctk.CTkButton(
            self,
            text="🗑️ Clear Selected Stage",
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            height=35,
            command=self._clear_stage
        )
        self.clear_btn.pack(pady=(0, 20))

    def _import_file(self):
        """Handle file import."""
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )

        if not file_path:
            return

        stage = self.stage_var.get()
        success, message, count = self.data_manager.import_from_docx(file_path, stage)

        self._update_status(f"[{stage}] {message}")

        if success:
            self._update_status(f"Current {stage} students: {self.data_manager.get_student_count(stage)}")
            messagebox.showinfo("Import Successful", message)
        else:
            messagebox.showerror("Import Failed", message)

    def _clear_stage(self):
        """Clear all students from the selected stage."""
        stage = self.stage_var.get()
        count = self.data_manager.get_student_count(stage)

        if count == 0:
            messagebox.showinfo("Info", f"No students in {stage} to clear.")
            return

        if messagebox.askyesno("Confirm", f"Clear all {count} students from {stage}?"):
            self.data_manager.clear_stage(stage)
            self._update_status(f"Cleared all students from {stage}")
            messagebox.showinfo("Success", f"Cleared {count} students from {stage}")

    def _update_status(self, message: str):
        """Update the status textbox."""
        self.status_text.configure(state="normal")
        self.status_text.insert("end", f"\n{message}")
        self.status_text.see("end")
        self.status_text.configure(state="disabled")

    def refresh(self):
        """Refresh the frame."""
        pass


class ViewDataFrame(ctk.CTkFrame):
    """Frame for viewing imported student data."""

    def __init__(self, master, data_manager: DataManager, **kwargs):
        super().__init__(master, **kwargs)

        self.data_manager = data_manager

        # Title
        self.title = ctk.CTkLabel(
            self,
            text="📊 View Student Data",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title.pack(pady=(40, 10))

        # Tab view for stages
        self.tabview = ctk.CTkTabview(self, width=500, height=400)
        self.tabview.pack(pady=20, padx=40, fill="both", expand=True)

        self.stage_tabs = {}
        self.stage_textboxes = {}

        for stage in self.data_manager.stages:
            tab = self.tabview.add(stage)
            self.stage_tabs[stage] = tab

            textbox = ctk.CTkTextbox(tab, font=ctk.CTkFont(size=12))
            textbox.pack(fill="both", expand=True, padx=10, pady=10)
            self.stage_textboxes[stage] = textbox

        # Summary Label
        self.summary_label = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=14)
        )
        self.summary_label.pack(pady=10)

        # Refresh Button
        self.refresh_btn = ctk.CTkButton(
            self,
            text="🔄 Refresh Data",
            command=self.refresh,
            height=35
        )
        self.refresh_btn.pack(pady=(0, 20))

    def refresh(self):
        """Refresh the displayed data."""
        total = 0

        for stage in self.data_manager.stages:
            students = self.data_manager.get_students(stage)
            textbox = self.stage_textboxes[stage]

            textbox.configure(state="normal")
            textbox.delete("1.0", "end")

            if students:
                for i, student in enumerate(students, 1):
                    textbox.insert("end", f"{i}. {student}\n")
                total += len(students)
            else:
                textbox.insert("1.0", "No students imported for this stage.")

            textbox.configure(state="disabled")

        self.summary_label.configure(
            text=f"Total Students: {total} | "
            f"1st: {self.data_manager.get_student_count('1st Stage')} | "
            f"2nd: {self.data_manager.get_student_count('2nd Stage')} | "
            f"3rd: {self.data_manager.get_student_count('3rd Stage')}"
        )


class GenerateFrame(ctk.CTkFrame):
    """Frame for generating seating arrangements."""

    def __init__(self, master, data_manager: DataManager, algorithm: SeatingAlgorithm, **kwargs):
        super().__init__(master, **kwargs)

        self.data_manager = data_manager
        self.algorithm = algorithm

        # Title
        self.title = ctk.CTkLabel(
            self,
            text="🪑 Generate Seating",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title.pack(pady=(40, 10))

        self.description = ctk.CTkLabel(
            self,
            text="Generate smart seating arrangements with cross-stage pairing",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        self.description.pack(pady=(0, 20))

        # Info Frame
        self.info_frame = ctk.CTkFrame(self)
        self.info_frame.pack(pady=10, padx=40, fill="x")

        self.info_label = ctk.CTkLabel(
            self.info_frame,
            text="ℹ️ Algorithm Rules:\n"
                 "• Students are paired from DIFFERENT stages\n"
                 "• Students within each stage are randomized\n"
                 "• Desks are organized into 3 columns (Left, Middle, Right)",
            font=ctk.CTkFont(size=12),
            justify="left"
        )
        self.info_label.pack(pady=15, padx=15)

        # Generate Button
        self.generate_btn = ctk.CTkButton(
            self,
            text="🎲 Generate Seating Arrangement",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            width=300,
            command=self._generate_seating
        )
        self.generate_btn.pack(pady=30)

        # Results Frame
        self.results_frame = ctk.CTkFrame(self)
        self.results_frame.pack(pady=10, padx=40, fill="both", expand=True)

        self.results_title = ctk.CTkLabel(
            self.results_frame,
            text="Generated Seating",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.results_title.pack(pady=(15, 10))

        self.results_text = ctk.CTkTextbox(
            self.results_frame,
            font=ctk.CTkFont(size=11, family="Courier")
        )
        self.results_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.results_text.insert("1.0", "Click 'Generate' to create seating arrangement...")
        self.results_text.configure(state="disabled")

        # Stats Label
        self.stats_label = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=12)
        )
        self.stats_label.pack(pady=(0, 20))

    def _generate_seating(self):
        """Generate the seating arrangement."""
        # Validate data
        is_valid, message = self.data_manager.validate_for_seating()
        if not is_valid:
            messagebox.showerror("Validation Error", message)
            return

        # Generate seating
        student_pools = self.data_manager.get_all_students()
        success, message = self.algorithm.generate_seating(student_pools)

        if not success:
            messagebox.showerror("Generation Error", message)
            return

        # Display results
        self._display_results()

        # Show stats
        stats = self.algorithm.get_statistics()
        self.stats_label.configure(
            text=f"✅ {stats['total_desks']} desks | "
            f"{stats['total_students']} students | "
            f"{stats['cross_stage_pairs']} cross-stage pairs"
        )

        if stats.get('same_stage_pairs', 0) > 0:
            messagebox.showwarning(
                "Warning",
                f"{stats['same_stage_pairs']} desk(s) have students from the same stage "
                "due to unbalanced numbers."
            )
        else:
            messagebox.showinfo("Success", message)

    def _display_results(self):
        """Display the generated seating arrangement."""
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")

        assignments = self.algorithm.assignments

        # Header
        header = f"{'Desk':<6} {'Column':<8} {'Student A':<25} {'Stage A':<12} {'Student B':<25} {'Stage B':<12}\n"
        self.results_text.insert("end", header)
        self.results_text.insert("end", "=" * 95 + "\n")

        # Data rows
        for a in assignments:
            student_b = a.student_b if a.student_b else "-"
            stage_b = a.student_b_stage if a.student_b_stage else "-"

            row = f"{a.desk_number:<6} {a.column:<8} {a.student_a:<25} {a.student_a_stage:<12} {student_b:<25} {stage_b:<12}\n"
            self.results_text.insert("end", row)

        self.results_text.configure(state="disabled")

    def refresh(self):
        """Refresh the frame."""
        pass


class ExportFrame(ctk.CTkFrame):
    """Frame for exporting seating arrangements to Word documents."""

    def __init__(self, master, data_manager: DataManager, algorithm: SeatingAlgorithm, exporter: WordExporter, **kwargs):
        super().__init__(master, **kwargs)

        self.data_manager = data_manager
        self.algorithm = algorithm
        self.exporter = exporter

        # Title
        self.title = ctk.CTkLabel(
            self,
            text="📄 Export to Word",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title.pack(pady=(40, 10))

        self.description = ctk.CTkLabel(
            self,
            text="Export the seating arrangement to a formatted Word document",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        self.description.pack(pady=(0, 30))

        # Options Frame
        self.options_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.options_frame.pack(pady=10, padx=40, fill="x")

        # Exam Title
        self.exam_title_label = ctk.CTkLabel(
            self.options_frame,
            text="Exam Title:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.exam_title_label.pack(anchor="w")

        self.exam_title_entry = ctk.CTkEntry(
            self.options_frame,
            placeholder_text="e.g., Final Exam Seating Arrangement",
            height=40,
            font=ctk.CTkFont(size=14)
        )
        self.exam_title_entry.pack(fill="x", pady=(5, 15))
        self.exam_title_entry.insert(0, "Exam Seating Arrangement")

        # Exam Date
        self.exam_date_label = ctk.CTkLabel(
            self.options_frame,
            text="Exam Date (optional):",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.exam_date_label.pack(anchor="w")

        self.exam_date_entry = ctk.CTkEntry(
            self.options_frame,
            placeholder_text="e.g., January 15, 2025",
            height=40,
            font=ctk.CTkFont(size=14)
        )
        self.exam_date_entry.pack(fill="x", pady=(5, 15))

        # Export Format
        self.format_label = ctk.CTkLabel(
            self.options_frame,
            text="Export Format:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.format_label.pack(anchor="w")

        self.format_var = ctk.StringVar(value="Standard Table")
        self.format_dropdown = ctk.CTkOptionMenu(
            self.options_frame,
            variable=self.format_var,
            values=["Standard Table", "Organized by Column"],
            width=200,
            height=40
        )
        self.format_dropdown.pack(anchor="w", pady=(5, 0))

        # Export Button
        self.export_btn = ctk.CTkButton(
            self,
            text="💾 Export to Word Document",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            width=300,
            command=self._export
        )
        self.export_btn.pack(pady=40)

        # Status
        self.status_label = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=13)
        )
        self.status_label.pack(pady=10)

    def _export(self):
        """Export the seating arrangement."""
        assignments = self.algorithm.assignments

        if not assignments:
            messagebox.showerror(
                "No Data",
                "No seating arrangement to export.\n"
                "Please generate seating first."
            )
            return

        # Get export options
        exam_title = self.exam_title_entry.get() or "Exam Seating Arrangement"
        exam_date = self.exam_date_entry.get() or None
        department = self.data_manager.department_name or "Unknown Department"

        # Get save path
        file_path = filedialog.asksaveasfilename(
            title="Save Word Document",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            initialfilename="exam_seating.docx"
        )

        if not file_path:
            return

        # Export based on format
        export_format = self.format_var.get()

        if export_format == "Organized by Column":
            success, message = self.exporter.export_by_column(
                assignments, department, file_path, exam_title, exam_date
            )
        else:
            success, message = self.exporter.export(
                assignments, department, file_path, exam_title, exam_date
            )

        if success:
            self.status_label.configure(text=f"✅ {message}", text_color="green")
            messagebox.showinfo("Export Successful", message)

            # Ask to open file
            if messagebox.askyesno("Open File", "Would you like to open the exported file?"):
                os.startfile(file_path) if os.name == 'nt' else os.system(f'xdg-open "{file_path}"')
        else:
            self.status_label.configure(text=f"❌ {message}", text_color="red")
            messagebox.showerror("Export Failed", message)

    def refresh(self):
        """Refresh the frame."""
        pass


class ExamSeatingApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("Exam Seating Management")
        self.geometry("1100x700")
        self.minsize(900, 600)

        # Set default theme
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # Initialize core components
        self.data_manager = DataManager()
        self.algorithm = SeatingAlgorithm()
        self.exporter = WordExporter()

        # Configure grid
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create Sidebar
        self.sidebar = SidebarFrame(
            self,
            nav_callback=self._navigate,
            width=200,
            corner_radius=0
        )
        self.sidebar.grid(row=0, column=0, sticky="nsw")

        # Create main content area
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        # Create content frames
        self.frames = {}
        self._create_frames()

        # Show home frame initially
        self._show_frame("home")

    def _create_frames(self):
        """Create all content frames."""
        self.frames["home"] = HomeFrame(
            self.main_frame,
            self.data_manager
        )

        self.frames["import"] = ImportFrame(
            self.main_frame,
            self.data_manager
        )

        self.frames["view"] = ViewDataFrame(
            self.main_frame,
            self.data_manager
        )

        self.frames["generate"] = GenerateFrame(
            self.main_frame,
            self.data_manager,
            self.algorithm
        )

        self.frames["export"] = ExportFrame(
            self.main_frame,
            self.data_manager,
            self.algorithm,
            self.exporter
        )

        # Place all frames in the same position
        for frame in self.frames.values():
            frame.grid(row=0, column=0, sticky="nsew")

    def _show_frame(self, key: str):
        """Show a specific frame."""
        frame = self.frames.get(key)
        if frame:
            frame.tkraise()
            # Refresh the frame if it has a refresh method
            if hasattr(frame, 'refresh'):
                frame.refresh()

    def _navigate(self, key: str):
        """Handle navigation from sidebar."""
        self._show_frame(key)


def main():
    """Main entry point."""
    app = ExamSeatingApp()
    app.mainloop()


if __name__ == "__main__":
    main()
