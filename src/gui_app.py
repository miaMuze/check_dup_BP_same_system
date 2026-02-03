"""
BP Duplicate Checker - Main GUI Application
============================================
This module contains the main Tkinter-based graphical user interface
for the BP Duplicate Checker application.

Features:
- File upload with drag-and-drop support indication
- Configurable ignore word list
- Progress bar for long operations
- Sortable results table
- Excel export functionality
"""

import os
import sys
import threading
from tkinter import (
    Tk, Frame, Label, Entry, Button, Text, Scrollbar,
    filedialog, messagebox, StringVar, IntVar, DoubleVar,
    HORIZONTAL, VERTICAL, BOTH, LEFT, RIGHT, TOP, BOTTOM,
    X, Y, END, W, E, N, S, DISABLED, NORMAL, CENTER
)
from tkinter import ttk
from typing import Optional, Dict, List
import queue

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.matching_engine import FuzzyMatcher, MatchResult
from src.excel_handler import ExcelHandler, create_example_input_file


class BPDuplicateCheckerApp:
    """
    Main application class for the BP Duplicate Checker GUI.
    """

    # Application constants
    APP_TITLE = "BP Duplicate Checker"
    APP_VERSION = "1.0.0"
    WINDOW_WIDTH = 1200
    WINDOW_HEIGHT = 700
    MIN_WIDTH = 900
    MIN_HEIGHT = 500

    # Default ignore words
    DEFAULT_IGNORE_WORDS = "Mrs, Ms, Mr, Dr, Prof, Company, Co, Ltd, LLC, Inc, Corp, Limited"

    def __init__(self, root: Tk):
        """
        Initialize the application.

        Args:
            root: The root Tkinter window
        """
        self.root = root
        self.setup_window()
        self.create_variables()
        self.create_widgets()
        self.setup_bindings()

        # Data storage
        self.loaded_data: List[Dict] = []
        self.matching_results: Dict = {}
        self.matcher: Optional[FuzzyMatcher] = None

        # Thread-safe queue for progress updates
        self.progress_queue = queue.Queue()

    def setup_window(self):
        """Configure the main window properties."""
        self.root.title(f"{self.APP_TITLE} v{self.APP_VERSION}")

        # Set window size and position (center on screen)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - self.WINDOW_WIDTH) // 2
        y = (screen_height - self.WINDOW_HEIGHT) // 2
        self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}+{x}+{y}")

        # Set minimum size
        self.root.minsize(self.MIN_WIDTH, self.MIN_HEIGHT)

        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use 'clam' theme for better appearance

        # Configure custom styles
        self.style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        self.style.configure('Status.TLabel', font=('Segoe UI', 9))
        self.style.configure('Action.TButton', font=('Segoe UI', 10))

    def create_variables(self):
        """Create Tkinter variables for data binding."""
        self.file_path_var = StringVar()
        self.ignore_words_var = StringVar(value=self.DEFAULT_IGNORE_WORDS)
        self.status_var = StringVar(value="Ready. Please upload an Excel file to begin.")
        self.progress_var = DoubleVar(value=0)
        self.min_score_var = IntVar(value=50)
        self.top_n_var = IntVar(value=3)

    def create_widgets(self):
        """Create all UI widgets."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # ===== Header Section =====
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=X, pady=(0, 10))

        title_label = ttk.Label(
            header_frame,
            text="Business Partner Duplicate Checker",
            style='Header.TLabel'
        )
        title_label.pack(side=LEFT)

        # Help button
        help_btn = ttk.Button(
            header_frame,
            text="Help",
            command=self.show_help,
            width=8
        )
        help_btn.pack(side=RIGHT)

        # ===== File Upload Section =====
        file_frame = ttk.LabelFrame(main_frame, text="Step 1: Upload Excel File", padding="10")
        file_frame.pack(fill=X, pady=(0, 10))

        # File path entry
        file_entry_frame = ttk.Frame(file_frame)
        file_entry_frame.pack(fill=X)

        self.file_entry = ttk.Entry(
            file_entry_frame,
            textvariable=self.file_path_var,
            state='readonly',
            width=80
        )
        self.file_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))

        browse_btn = ttk.Button(
            file_entry_frame,
            text="Browse...",
            command=self.browse_file,
            style='Action.TButton'
        )
        browse_btn.pack(side=LEFT)

        # File info label
        self.file_info_label = ttk.Label(
            file_frame,
            text="Supported format: Excel (.xlsx)",
            foreground="gray"
        )
        self.file_info_label.pack(anchor=W, pady=(5, 0))

        # ===== Configuration Section =====
        config_frame = ttk.LabelFrame(main_frame, text="Step 2: Configure Matching", padding="10")
        config_frame.pack(fill=X, pady=(0, 10))

        # Ignore words
        ignore_frame = ttk.Frame(config_frame)
        ignore_frame.pack(fill=X, pady=(0, 10))

        ignore_label = ttk.Label(ignore_frame, text="Ignore Words (comma-separated):")
        ignore_label.pack(anchor=W)

        self.ignore_entry = ttk.Entry(
            ignore_frame,
            textvariable=self.ignore_words_var,
            width=100
        )
        self.ignore_entry.pack(fill=X, pady=(5, 0))

        # Options row
        options_frame = ttk.Frame(config_frame)
        options_frame.pack(fill=X)

        # Minimum score
        min_score_label = ttk.Label(options_frame, text="Minimum Score (%):")
        min_score_label.pack(side=LEFT)

        self.min_score_spin = ttk.Spinbox(
            options_frame,
            from_=0,
            to=100,
            width=5,
            textvariable=self.min_score_var
        )
        self.min_score_spin.pack(side=LEFT, padx=(5, 20))

        # Top N matches
        top_n_label = ttk.Label(options_frame, text="Top N Matches:")
        top_n_label.pack(side=LEFT)

        self.top_n_spin = ttk.Spinbox(
            options_frame,
            from_=1,
            to=10,
            width=5,
            textvariable=self.top_n_var
        )
        self.top_n_spin.pack(side=LEFT, padx=(5, 20))

        # Run button
        self.run_btn = ttk.Button(
            options_frame,
            text="Run Matching",
            command=self.run_matching,
            style='Action.TButton'
        )
        self.run_btn.pack(side=RIGHT)

        # ===== Results Section =====
        results_frame = ttk.LabelFrame(main_frame, text="Step 3: Review Results", padding="10")
        results_frame.pack(fill=BOTH, expand=True, pady=(0, 10))

        # Results Treeview
        self.create_results_table(results_frame)

        # ===== Action Buttons =====
        action_frame = ttk.Frame(results_frame)
        action_frame.pack(fill=X, pady=(10, 0))

        self.export_btn = ttk.Button(
            action_frame,
            text="Export to Excel",
            command=self.export_results,
            state=DISABLED,
            style='Action.TButton'
        )
        self.export_btn.pack(side=RIGHT, padx=(10, 0))

        self.clear_btn = ttk.Button(
            action_frame,
            text="Clear Results",
            command=self.clear_results,
            state=DISABLED,
            style='Action.TButton'
        )
        self.clear_btn.pack(side=RIGHT)

        # Result count label
        self.result_count_label = ttk.Label(action_frame, text="")
        self.result_count_label.pack(side=LEFT)

        # ===== Status Bar =====
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=X)

        self.progress_bar = ttk.Progressbar(
            status_frame,
            mode='determinate',
            variable=self.progress_var,
            length=200
        )
        self.progress_bar.pack(side=RIGHT)

        status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style='Status.TLabel'
        )
        status_label.pack(side=LEFT, fill=X, expand=True)

    def create_results_table(self, parent):
        """
        Create the results treeview table with scrollbars.

        Args:
            parent: Parent frame for the table
        """
        # Create frame for treeview and scrollbars
        table_frame = ttk.Frame(parent)
        table_frame.pack(fill=BOTH, expand=True)

        # Define columns
        columns = (
            'source_bp', 'source_name1', 'source_name2',
            'rank', 'match_bp', 'match_name1', 'match_name2',
            'score', 'confidence'
        )

        # Create treeview
        self.results_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show='headings',
            selectmode='browse'
        )

        # Configure columns
        column_config = {
            'source_bp': ('Source BP#', 100, CENTER),
            'source_name1': ('Source Name1', 150, W),
            'source_name2': ('Source Name2', 120, W),
            'rank': ('Rank', 50, CENTER),
            'match_bp': ('Match BP#', 100, CENTER),
            'match_name1': ('Match Name1', 150, W),
            'match_name2': ('Match Name2', 120, W),
            'score': ('Score', 70, CENTER),
            'confidence': ('Confidence', 90, CENTER)
        }

        for col, (heading, width, anchor) in column_config.items():
            self.results_tree.heading(col, text=heading, command=lambda c=col: self.sort_column(c))
            self.results_tree.column(col, width=width, anchor=anchor, minwidth=50)

        # Scrollbars
        v_scroll = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.results_tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient=HORIZONTAL, command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        # Grid layout
        self.results_tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Configure row tags for conditional formatting
        self.results_tree.tag_configure('high', background='#ffcccc')
        self.results_tree.tag_configure('medium', background='#fff3cd')
        self.results_tree.tag_configure('low', background='#ffffff')

    def setup_bindings(self):
        """Set up keyboard and event bindings."""
        self.root.bind('<Control-o>', lambda e: self.browse_file())
        self.root.bind('<Control-e>', lambda e: self.export_results())
        self.root.bind('<F5>', lambda e: self.run_matching())

    def browse_file(self):
        """Open file dialog to select Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("All Files", "*.*")
            ]
        )

        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path: str):
        """
        Load and validate the selected Excel file.

        Args:
            file_path: Path to the Excel file
        """
        self.update_status("Validating file...")

        # Validate file
        is_valid, message = ExcelHandler.validate_file(file_path)

        if not is_valid:
            messagebox.showerror("Validation Error", message)
            self.update_status("File validation failed")
            return

        # Load data
        self.loaded_data, load_message = ExcelHandler.load_data(file_path)

        if not self.loaded_data:
            messagebox.showerror("Load Error", load_message)
            self.update_status("Failed to load data")
            return

        # Update UI
        self.file_path_var.set(file_path)
        self.file_info_label.config(
            text=f"Loaded: {len(self.loaded_data)} records",
            foreground="green"
        )
        self.update_status(load_message)

        # Clear previous results
        self.clear_results()

    def run_matching(self):
        """Start the fuzzy matching process in a background thread."""
        if not self.loaded_data:
            messagebox.showwarning("No Data", "Please upload an Excel file first.")
            return

        # Parse ignore words
        ignore_text = self.ignore_words_var.get()
        ignore_words = [word.strip() for word in ignore_text.split(',') if word.strip()]

        # Get options
        min_score = self.min_score_var.get()
        top_n = self.top_n_var.get()

        # Disable UI during processing
        self.set_ui_state(False)
        self.progress_var.set(0)
        self.update_status("Running fuzzy matching...")

        # Run matching in background thread
        thread = threading.Thread(
            target=self.matching_worker,
            args=(ignore_words, min_score, top_n),
            daemon=True
        )
        thread.start()

        # Start progress monitoring
        self.root.after(100, self.check_progress)

    def matching_worker(self, ignore_words: List[str], min_score: int, top_n: int):
        """
        Background worker for fuzzy matching.

        Args:
            ignore_words: List of words to ignore
            min_score: Minimum similarity score threshold
            top_n: Number of top matches to return
        """
        try:
            # Create matcher with ignore words
            self.matcher = FuzzyMatcher(ignore_words)
            self.matcher.load_records(self.loaded_data)

            # Progress callback
            def progress_callback(current, total):
                progress = (current / total) * 100
                self.progress_queue.put(('progress', progress))

            # Run matching
            self.matching_results = self.matcher.find_matches(
                top_n=top_n,
                min_score=min_score,
                progress_callback=progress_callback
            )

            # Signal completion
            self.progress_queue.put(('complete', None))

        except Exception as e:
            self.progress_queue.put(('error', str(e)))

    def check_progress(self):
        """Check progress queue and update UI."""
        try:
            while True:
                msg_type, data = self.progress_queue.get_nowait()

                if msg_type == 'progress':
                    self.progress_var.set(data)

                elif msg_type == 'complete':
                    self.progress_var.set(100)
                    self.display_results()
                    self.set_ui_state(True)
                    return

                elif msg_type == 'error':
                    messagebox.showerror("Matching Error", f"An error occurred: {data}")
                    self.set_ui_state(True)
                    self.update_status("Matching failed")
                    return

        except queue.Empty:
            pass

        # Continue checking
        self.root.after(100, self.check_progress)

    def display_results(self):
        """Display matching results in the treeview."""
        # Clear existing results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)

        # Count total matches
        total_matches = 0
        records_with_matches = 0

        # Insert results
        for bp_number, matches in self.matching_results.items():
            if not matches:
                continue

            records_with_matches += 1

            for rank, match in enumerate(matches, 1):
                total_matches += 1
                score = match.similarity_score

                # Determine confidence level and tag
                if score >= 80:
                    confidence = "High"
                    tag = 'high'
                elif score >= 60:
                    confidence = "Medium"
                    tag = 'medium'
                else:
                    confidence = "Low"
                    tag = 'low'

                # Insert row
                self.results_tree.insert('', 'end', values=(
                    match.source_bp.bp_number,
                    match.source_bp.name1,
                    match.source_bp.name2,
                    rank,
                    match.match_bp.bp_number,
                    match.match_bp.name1,
                    match.match_bp.name2,
                    f"{score:.1f}%",
                    confidence
                ), tags=(tag,))

        # Update UI
        self.result_count_label.config(
            text=f"Found {total_matches} potential matches across {records_with_matches} records"
        )
        self.export_btn.config(state=NORMAL if total_matches > 0 else DISABLED)
        self.clear_btn.config(state=NORMAL)
        self.update_status(f"Matching complete. Found {total_matches} potential duplicates.")

    def sort_column(self, col: str):
        """
        Sort treeview by column.

        Args:
            col: Column identifier to sort by
        """
        # Get all items
        items = [(self.results_tree.set(item, col), item) for item in self.results_tree.get_children()]

        # Determine sort order (toggle)
        reverse = getattr(self, f'_sort_{col}_reverse', False)

        # Sort - handle numeric columns specially
        if col in ('rank', 'score'):
            # Extract numeric value for sorting
            def sort_key(x):
                val = x[0].replace('%', '').strip()
                try:
                    return float(val)
                except ValueError:
                    return 0
            items.sort(key=sort_key, reverse=reverse)
        else:
            items.sort(key=lambda x: x[0].lower(), reverse=reverse)

        # Rearrange items
        for index, (val, item) in enumerate(items):
            self.results_tree.move(item, '', index)

        # Toggle sort order for next time
        setattr(self, f'_sort_{col}_reverse', not reverse)

    def export_results(self):
        """Export results to Excel file."""
        if not self.matching_results:
            messagebox.showwarning("No Results", "No results to export. Run matching first.")
            return

        # Get output file path
        output_path = filedialog.asksaveasfilename(
            title="Save Results",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"BP_Duplicate_Check_Results.xlsx"
        )

        if not output_path:
            return

        self.update_status("Exporting results...")

        # Get summary statistics
        summary_stats = self.matcher.get_summary_stats(self.matching_results) if self.matcher else None

        # Export
        success, message = ExcelHandler.export_results(
            self.matching_results,
            output_path,
            summary_stats
        )

        if success:
            messagebox.showinfo("Export Complete", message)
            self.update_status("Results exported successfully")
        else:
            messagebox.showerror("Export Error", message)
            self.update_status("Export failed")

    def clear_results(self):
        """Clear all results from the display."""
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)

        self.matching_results = {}
        self.result_count_label.config(text="")
        self.export_btn.config(state=DISABLED)
        self.clear_btn.config(state=DISABLED)
        self.progress_var.set(0)
        self.update_status("Results cleared")

    def set_ui_state(self, enabled: bool):
        """
        Enable or disable UI elements during processing.

        Args:
            enabled: True to enable, False to disable
        """
        state = NORMAL if enabled else DISABLED
        self.run_btn.config(state=state)
        self.ignore_entry.config(state=state)
        self.min_score_spin.config(state=state)
        self.top_n_spin.config(state=state)

    def update_status(self, message: str):
        """
        Update the status bar message.

        Args:
            message: Status message to display
        """
        self.status_var.set(message)
        self.root.update_idletasks()

    def show_help(self):
        """Display help dialog."""
        help_text = """BP Duplicate Checker - Help

USAGE:
1. Click 'Browse' to select an Excel file (.xlsx)
   Required columns: BP_Number, Name1, Name2

2. Configure matching options:
   - Ignore Words: Words to exclude from comparison
   - Minimum Score: Only show matches above this threshold
   - Top N Matches: Number of matches to show per BP

3. Click 'Run Matching' to find duplicates

4. Review results in the table:
   - Red rows: High confidence matches (≥80%)
   - Yellow rows: Medium confidence matches (60-79%)
   - White rows: Low confidence matches (<60%)
   - Click column headers to sort

5. Click 'Export to Excel' to save results

KEYBOARD SHORTCUTS:
- Ctrl+O: Open file
- Ctrl+E: Export results
- F5: Run matching

TIPS:
- Add common titles (Mr, Mrs) to ignore words
- Add business suffixes (Ltd, Inc) to ignore words
- Higher minimum score = fewer but better matches
"""
        messagebox.showinfo("Help", help_text)


def main():
    """Main entry point for the application."""
    root = Tk()
    app = BPDuplicateCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
