# range_paster_gui.py
# Requires: pip install openpyxl
import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Tuple, Optional, List, Dict, Any
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ----------------------------- Module logger -----------------------------
LOGGER = None  # set to a callable(str) by the GUI so helpers can log to the Output Log

def set_logger(fn):
    global LOGGER
    LOGGER = fn

def log_debug(msg: str):
    if LOGGER:
        LOGGER(msg)
    else:
        print(msg)

# ----------------------------- Core helpers -----------------------------

def last_nonempty_col_in_row_range(ws: Worksheet, start_col: int, min_row: int, max_row: int) -> int:
    """
    Scan columns from start_col through ws.max_column, and within the given row range [min_row..max_row],
    return the last column index that contains any non-empty cell.
    """
    last = start_col
    max_scan_col = ws.max_column
    for c in range(start_col, max_scan_col + 1):
        for r in range(min_row, max_row + 1):
            if cell_is_nonempty(ws, r, c):
                last = c
                break
    log_debug(f"[calc] last_nonempty_col_in_row_range rows={min_row}..{max_row} start_col={start_col} -> last_col={last}")
    return last


def format_range_a1(min_c: int, min_r: int, max_c: int, max_r: int, sheet: Optional[str]) -> str:
    start = f"{get_column_letter(min_c)}{min_r}"
    end   = f"{get_column_letter(max_c)}{max_r}"
    core  = start if (min_c == max_c and min_r == max_r) else f"{start}:{end}"
    return f"{sheet}!{core}" if sheet else core


def parse_sheet_and_ref(token: str) -> Tuple[Optional[str], str]:
    token = token.strip()
    if "!" in token:
        sheet, ref = token.split("!", 1)
        res = (sheet.strip(), ref.strip())
        log_debug(f"[calc] parse_sheet_and_ref '{token}' -> sheet='{res[0]}', ref='{res[1]}'")
        return res
    log_debug(f"[calc] parse_sheet_and_ref '{token}' -> active sheet, ref='{token}'")
    return None, token

def parse_cell(ref: str) -> Tuple[int, int]:
    from openpyxl.utils.cell import coordinate_to_tuple
    row, col = coordinate_to_tuple(ref)
    log_debug(f"[calc] parse_cell '{ref}' -> col={col}, row={row}")
    return col, row

def parse_range_standard(ref: str) -> Tuple[int, int, int, int]:
    from openpyxl.utils.cell import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    log_debug(f"[calc] parse_range_standard '{ref}' -> ({min_col},{min_row})..({max_col},{max_row})")
    return min_col, min_row, max_col, max_row

def get_sheet(wb, name: Optional[str]) -> Worksheet:
    if name is None:
        ws = wb.active
        log_debug(f"[calc] get_sheet active -> '{ws.title}'")
        return ws
    if name in wb.sheetnames:
        log_debug(f"[calc] get_sheet found -> '{name}'")
        return wb[name]
    msg = f'Sheet "{name}" not found in workbook.'
    log_debug(f"[calc] get_sheet error -> {msg}")
    raise ValueError(msg)

def cell_is_nonempty(ws: Worksheet, row: int, col: int) -> bool:
    v = ws.cell(row=row, column=col).value
    return v is not None and (not isinstance(v, str) or v.strip() != "")

def last_nonempty_col_on_row(ws: Worksheet, row: int, start_col: int) -> int:
    """
    Scan columns on a single row from start_col through ws.max_column and return the last non-empty column index.
    """
    last = start_col
    max_scan_col = ws.max_column
    for c in range(start_col, max_scan_col + 1):
        if cell_is_nonempty(ws, row, c):
            last = c
    log_debug(f"[calc] last_nonempty_col_on_row row={row} from_col={start_col} -> last_col={last}")
    return last


def last_nonempty_row_on_col(ws: Worksheet, col: int, start_row: int) -> int:
    last = start_row
    for r in range(start_row, ws.max_row + 1):
        if cell_is_nonempty(ws, r, col):
            last = r
    log_debug(f"[calc] last_nonempty_row_on_col col={col} from_row={start_row} -> last_row={last}")
    return last

def first_nonempty_row_on_col(ws: Worksheet, col: int, start_row: int = 1) -> int:
    """
    Scan downward from start_row on a single column and return the first row index
    that has a non-empty cell. If none found, return start_row.
    """
    for r in range(start_row, ws.max_row + 1):
        if cell_is_nonempty(ws, r, col):
            log_debug(f"[calc] first_nonempty_row_on_col col={col} -> first_row={r}")
            return r
    log_debug(f"[calc] first_nonempty_row_on_col col={col} -> none found, default={start_row}")
    return start_row


def last_nonempty_row_in_col_range(ws: Worksheet, min_col: int, max_col: int, start_row: int) -> int:
    """
    Within columns [min_col..max_col], find the last row at or below start_row
    where ANY column in that span has a non-empty cell.
    """
    last = start_row
    for r in range(start_row, ws.max_row + 1):
        any_nonempty = False
        for c in range(min_col, max_col + 1):
            if cell_is_nonempty(ws, r, c):
                any_nonempty = True
                break
        if any_nonempty:
            last = r
    log_debug(f"[calc] last_nonempty_row_in_col_range cols={min_col}..{max_col} from_row={start_row} -> last_row={last}")
    return last


def split_col_row(token: str) -> Tuple[Optional[str], Optional[str]]:
    token = token.strip()
    if token == "??":
        log_debug(f"[calc] split_col_row '{token}' -> col='?', row='?'")
        return "?", "?"
    i = 0
    while i < len(token) and token[i].isalpha():
        i += 1
    col_part = token[:i] if i > 0 else None
    row_part = token[i:] if i < len(token) else None
    if col_part is None and token and token[0] == "?":
        col_part = "?"
        row_part = token[1:] if len(token) > 1 else None
    if row_part is None and token and token[-1] == "?":
        row_part = "?"
    log_debug(f"[calc] split_col_row '{token}' -> col='{col_part}', row='{row_part}'")
    return col_part, row_part

def parse_range_with_wildcards_basic(ws: Worksheet, ref: str) -> Tuple[int, int, int, int]:
    """
    Original wildcard handler: A23:B? or A23:?56 or A23:??
    Uses row/col expansion based on existing non-empty content.
    """
    left, right = [x.strip() for x in ref.split(":", 1)]
    start_col, start_row = parse_cell(left)
    end_col_tok, end_row_tok = split_col_row(right)

    if end_col_tok == "?":
        end_col = last_nonempty_col_on_row(ws, start_row, start_col)
    elif end_col_tok is None:
        end_col = start_col
    else:
        end_col = column_index_from_string(end_col_tok)

    if end_row_tok == "?":
        end_row = last_nonempty_row_on_col(ws, start_col, start_row)
    elif end_row_tok is None:
        end_row = start_row
    else:
        try:
            end_row = int(end_row_tok)
        except ValueError:
            raise ValueError(f'Invalid row in "{right}"')

    min_c, max_c = sorted((start_col, end_col))
    min_r, max_r = sorted((start_row, end_row))
    log_debug(f"[calc] parse_range_with_wildcards_basic '{ref}' -> ({min_c},{min_r})..({max_c},{max_r})")
    return min_c, min_r, max_c, max_r

def copy_values(src_ws: Worksheet, dst_ws: Worksheet,
                src_bounds: Tuple[int, int, int, int],
                dst_row: int, dst_col: int) -> Tuple[int, int]:
    min_c, min_r, max_c, max_r = src_bounds
    width = max_c - min_c + 1
    height = max_r - min_r + 1
    log_debug(f"[calc] copy_values bounds=({min_c},{min_r})..({max_c},{max_r}) -> width={width}, height={height}, paste_to=({dst_col},{dst_row})")

    for r_off in range(height):
        for c_off in range(width):
            v = src_ws.cell(row=min_r + r_off, column=min_c + c_off).value
            dst_ws.cell(row=dst_row + r_off, column=dst_col + c_off).value = v

    return width, height

def normalize_value(v: Any) -> str:
    """Stringify and strip to compare values robustly."""
    if v is None:
        return ""
    return str(v).strip()

def find_row_by_value(ws: Worksheet, search_value: Any) -> Optional[int]:
    """
    Search for search_value across columns 1..ws.max_column, but skip columns that are entirely empty.
    Returns the 1-based row index of the first match (top-down), or None if not found.
    """
    target = normalize_value(search_value)
    if not target:
        log_debug(f"[search] find_row_by_value empty target -> None")
        return None

    max_col = ws.max_column
    log_debug(f"[search] find_row_by_value target='{target}' scanning cols 1..{max_col}, rows 1..{ws.max_row}")

    for col in range(1, max_col + 1):
        # Skip columns that are completely empty to avoid wasted scans
        any_nonempty = False
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=col).value is not None:
                any_nonempty = True
                break
        if not any_nonempty:
            continue

        for r in range(1, ws.max_row + 1):
            if normalize_value(ws.cell(row=r, column=col).value) == target:
                log_debug(f"[search] find_row_by_value found at row={r}, col={col}")
                return r
    log_debug(f"[search] find_row_by_value not found -> None")
    return None


# ----------------------------- UI pieces -----------------------------

class BaseCellFrame(ttk.Frame):
    """
    A 'track' that owns a Base Cell plus optional Start/End Reference cells.
    """
    def __init__(self, parent, index_changed_cb, remove_cb):
        super().__init__(parent)
        self.index_changed_cb = index_changed_cb
        self.remove_cb = remove_cb

        # Base cell
        row0 = ttk.Frame(self)
        row0.pack(fill="x", pady=2)
        ttk.Label(row0, text="Base Cell:").pack(side="left")
        self.base_cell_entry = ttk.Entry(row0, width=12)
        self.base_cell_entry.pack(side="left", padx=(5, 10))
        self.base_cell_entry.insert(0, "A1")

        # Optional Start/End references
        ttk.Label(row0, text="Start Row Ref:").pack(side="left")
        self.start_ref_entry = ttk.Entry(row0, width=12)
        self.start_ref_entry.pack(side="left", padx=(5, 10))

        ttk.Label(row0, text="End Row Ref:").pack(side="left")
        self.end_ref_entry = ttk.Entry(row0, width=12)
        self.end_ref_entry.pack(side="left", padx=(5, 10))

        ttk.Button(row0, text="Remove", command=self._remove_self).pack(side="left", padx=(5, 0))

        # Help
        help_text = "You can include sheet: e.g. Sheet1!C13"
        ttk.Label(self, text=help_text, foreground="gray", font=("TkDefaultFont", 8)).pack(anchor="w", padx=(0, 0))

        # When base cell changes, notify so dropdowns can refresh labels if needed
        self.base_cell_entry.bind("<FocusOut>", lambda e: self.index_changed_cb())

    def _remove_self(self):
        self.remove_cb(self)

    def get_label(self) -> str:
        # For dropdown display; include the entered base cell string
        val = self.base_cell_entry.get().strip() or "(unset)"
        return val

    def get_data(self) -> Dict[str, str]:
        return {
            "base_cell": self.base_cell_entry.get().strip(),
            "start_ref": self.start_ref_entry.get().strip(),
            "end_ref": self.end_ref_entry.get().strip(),
        }

    def set_data(self, data: Dict[str, str]):
        self.base_cell_entry.delete(0, tk.END)
        self.base_cell_entry.insert(0, data.get("base_cell", "A1"))
        self.start_ref_entry.delete(0, tk.END)
        self.start_ref_entry.insert(0, data.get("start_ref", ""))
        self.end_ref_entry.delete(0, tk.END)
        self.end_ref_entry.insert(0, data.get("end_ref", ""))

class SourceFileFrame(ttk.Frame):
    def __init__(self, parent, on_remove_callback, base_cells_provider):
        """
        base_cells_provider(): returns list of labels for base cells
        """
        super().__init__(parent)
        self.on_remove_callback = on_remove_callback
        self.base_cells_provider = base_cells_provider
        self.file_path = ""

        # NEW: Title row
        self.title_frame = ttk.Frame(self)
        self.title_frame.pack(fill="x", padx=5, pady=(0, 2))
        ttk.Label(self.title_frame, text="Title:").pack(side="left")
        self.title_entry = ttk.Entry(self.title_frame, width=40)
        self.title_entry.pack(side="left", padx=(5, 0))

        # File selection
        self.file_frame = ttk.Frame(self)
        self.file_frame.pack(fill="x", padx=5, pady=2)

        ttk.Label(self.file_frame, text="Source File:").pack(side="left")
        self.file_label = ttk.Label(self.file_frame, text="No file selected",
                                    foreground="gray", width=40, anchor="w")
        self.file_label.pack(side="left", padx=(5, 0))

        ttk.Button(self.file_frame, text="Browse",
                   command=self.browse_file).pack(side="left", padx=(5, 0))
        ttk.Button(self.file_frame, text="Remove",
                   command=self.remove_self).pack(side="left", padx=(5, 0))

        # Range + Base-cell dropdown
        row2 = ttk.Frame(self)
        row2.pack(fill="x", padx=5, pady=2)

        ttk.Label(row2, text="Source Range:").pack(side="left")
        self.range_entry = ttk.Entry(row2, width=22)
        self.range_entry.pack(side="left", padx=(5, 10))

        ttk.Label(row2, text="Track (Base Cell):").pack(side="left")
        self.track_var = tk.StringVar()
        self.track_combo = ttk.Combobox(row2, textvariable=self.track_var, width=16, state="readonly")
        self.track_combo.pack(side="left", padx=(5, 0))
        self.refresh_tracks()

        # Help text
        help_text = "Examples: A1:C10, Sheet1!B5:D15, N?:B?, A23:B?, A23:?56, A23:??"
        ttk.Label(self, text=help_text, foreground="gray", font=("TkDefaultFont", 8)).pack(anchor="w", padx=5)

    def refresh_tracks(self):
        labels = self.base_cells_provider()
        if not labels:
            labels = ["(none)"]
        self.track_combo["values"] = labels
        if labels:
            self.track_combo.current(0)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Source Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=filename, foreground="black")

    def remove_self(self):
        self.on_remove_callback(self)

    def get_data(self):
        return {
            'title': self.title_entry.get().strip(),
            'file_path': self.file_path,
            'range': self.range_entry.get().strip(),
            'track_label': self.track_var.get(),
        }

    def set_data(self, data: Dict[str, str]):
        self.file_path = data.get("file_path", "")
        self.file_label.config(
            text=(os.path.basename(self.file_path) if self.file_path else "No file selected"),
            foreground=("black" if self.file_path else "gray")
        )

        # NEW: restore title
        self.title_entry.delete(0, tk.END)
        self.title_entry.insert(0, data.get("title", ""))

        self.range_entry.delete(0, tk.END)
        self.range_entry.insert(0, data.get("range", ""))

        # track selection will be applied by caller after all base cells exist

    def is_valid(self):
        return bool(self.file_path and self.range_entry.get().strip())

# ----------------------------- Main GUI -----------------------------

class RangePasterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Range Paster")
        self.root.geometry("980x800")

        # Resizable
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.base_file_path = ""
        self.base_frames: List[BaseCellFrame] = []
        self.source_frames: List[SourceFileFrame] = []

        self.create_widgets()
        # Wire up module logger to the GUI output log
        set_logger(self.log)

    # ----- UI scaffolding -----

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        for r in (2, 4):
            main_frame.grid_rowconfigure(r, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        title_label = ttk.Label(main_frame, text="Excel Range Paster", font=("TkDefaultFont", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 12), sticky="w")

        # Base file
        base_frame = ttk.LabelFrame(main_frame, text="Base File", padding=10)
        base_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        base_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(base_frame, text="Base File:").grid(row=0, column=0, sticky="w")
        self.base_file_label = ttk.Label(base_frame, text="No file selected", foreground="gray", anchor="w")
        self.base_file_label.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        ttk.Button(base_frame, text="Browse", command=self.browse_base_file).grid(row=0, column=2, padx=(5, 0))

        # Base cells (tracks)
        tracks_frame = ttk.LabelFrame(main_frame, text="Base Cells (Tracks)", padding=8)
        tracks_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        tracks_frame.grid_rowconfigure(0, weight=1)
        tracks_frame.grid_columnconfigure(0, weight=1)

        # scrollable
        self.tracks_canvas = tk.Canvas(tracks_frame)
        self.tracks_scroll = ttk.Scrollbar(tracks_frame, orient="vertical", command=self.tracks_canvas.yview)
        self.tracks_container = ttk.Frame(self.tracks_canvas)
        self.tracks_container.bind("<Configure>", lambda e: self.tracks_canvas.configure(scrollregion=self.tracks_canvas.bbox("all")))
        self.tracks_canvas.create_window((0, 0), window=self.tracks_container, anchor="nw")
        self.tracks_canvas.configure(yscrollcommand=self.tracks_scroll.set)
        self.tracks_canvas.grid(row=0, column=0, sticky="nsew")
        self.tracks_scroll.grid(row=0, column=1, sticky="ns")

        add_track_frame = ttk.Frame(tracks_frame)
        add_track_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky="w")
        ttk.Button(add_track_frame, text="+ Add Base Cell", command=self.add_base_cell).pack(side="left")

        # Source files
        src_frame = ttk.LabelFrame(main_frame, text="Source Files", padding=8)
        src_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 10))
        src_frame.grid_rowconfigure(0, weight=1)
        src_frame.grid_columnconfigure(0, weight=1)

        self.src_canvas = tk.Canvas(src_frame)
        self.src_scroll = ttk.Scrollbar(src_frame, orient="vertical", command=self.src_canvas.yview)
        self.src_container = ttk.Frame(self.src_canvas)
        self.src_container.bind("<Configure>", lambda e: self.src_canvas.configure(scrollregion=self.src_canvas.bbox("all")))
        self.src_canvas.create_window((0, 0), window=self.src_container, anchor="nw")
        self.src_canvas.configure(yscrollcommand=self.src_scroll.set)
        self.src_canvas.grid(row=0, column=0, sticky="nsew")
        self.src_scroll.grid(row=0, column=1, sticky="ns")

        add_src_frame = ttk.Frame(src_frame)
        add_src_frame.grid(row=1, column=0, columnspan=2, pady=5)
        ttk.Button(add_src_frame, text="+ Add Source File", command=self.add_source_file).pack()

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, pady=(0, 10), sticky="w")

        ttk.Button(button_frame, text="Process Ranges", command=self.process_ranges).pack(side="left", padx=(0, 10))
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side="left", padx=(0, 20))

        ttk.Button(button_frame, text="Save Workspace", command=self.save_workspace).pack(side="left", padx=(0, 10))
        ttk.Button(button_frame, text="Load Workspace", command=self.load_workspace).pack(side="left")

        # Output log
        output_frame = ttk.LabelFrame(main_frame, text="Output Log", padding=5)
        output_frame.grid(row=6, column=0, sticky="ew")
        output_frame.grid_columnconfigure(0, weight=1)
        self.output_text = scrolledtext.ScrolledText(output_frame, height=8, width=80)
        self.output_text.grid(row=0, column=0, sticky="ew")

        # Add initial one track and one source
        self.add_base_cell()     # default A1 with optional refs blank
        self.add_source_file()

    # ----- Track utilities -----

    def get_track_labels(self) -> List[str]:
        return [bf.get_label() for bf in self.base_frames]

    def on_track_label_changed(self):
        # refresh each source dropdown to reflect updated labels
        labels = self.get_track_labels()
        for sf in self.source_frames:
            current = sf.track_var.get()
            sf.refresh_tracks()
            # try to keep same selection if still present
            if current in sf.track_combo["values"]:
                sf.track_var.set(current)

    def add_base_cell(self):
        def on_changed():
            self.on_track_label_changed()

        frame = BaseCellFrame(self.tracks_container, index_changed_cb=on_changed, remove_cb=self.remove_base_cell)
        frame.pack(fill="x", pady=3)
        self.base_frames.append(frame)
        self.root.update_idletasks()
        self.on_track_label_changed()

    def remove_base_cell(self, frame: BaseCellFrame):
        if len(self.base_frames) <= 1:
            messagebox.showinfo("Info", "At least one Base Cell is required.")
            return
        frame.destroy()
        self.base_frames.remove(frame)
        self.root.update_idletasks()
        self.on_track_label_changed()

    # ----- Source utilities -----

    def add_source_file(self):
        frame = SourceFileFrame(self.src_container, self.remove_source_file, base_cells_provider=self.get_track_labels)
        frame.pack(fill="x", pady=2)
        self.source_frames.append(frame)
        self.root.update_idletasks()

    def remove_source_file(self, frame):
        if len(self.source_frames) > 1:
            frame.destroy()
            self.source_frames.remove(frame)
            self.root.update_idletasks()

    # ----- Logging -----

    def log(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update()

    # ----- File browsing / reset -----

    def browse_base_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Base Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.base_file_path = file_path
            filename = os.path.basename(file_path)
            self.base_file_label.config(text=filename, foreground="black")

    def clear_all(self):
        self.base_file_path = ""
        self.base_file_label.config(text="No file selected", foreground="gray")
        # reset tracks
        for f in self.base_frames:
            f.destroy()
        self.base_frames.clear()
        self.add_base_cell()

        # reset sources (keep one)
        for f in self.source_frames[1:]:
            f.destroy()
        self.source_frames = self.source_frames[:1]
        sf = self.source_frames[0]
        sf.file_path = ""
        sf.file_label.config(text="No file selected", foreground="gray")
        sf.range_entry.delete(0, tk.END)
        sf.refresh_tracks()
        sf.title_entry.delete(0, tk.END)

        self.output_text.delete(1.0, tk.END)
        self.log("Cleared all inputs.")

    # ----- Validation -----

    def validate_inputs(self) -> bool:
        if not self.base_file_path:
            messagebox.showerror("Error", "Please select a base file.")
            return False
        if not os.path.isfile(self.base_file_path):
            messagebox.showerror("Error", f"Base file not found: {self.base_file_path}")
            return False

        if not self.base_frames:
            messagebox.showerror("Error", "Please add at least one Base Cell.")
            return False

        # Validate base cell formats
        for i, bf in enumerate(self.base_frames, start=1):
            data = bf.get_data()
            if not data["base_cell"]:
                messagebox.showerror("Error", f"Base Cell missing for track #{i}.")
                return False

        valid_sources = [f for f in self.source_frames if f.is_valid()]
        if not valid_sources:
            messagebox.showerror("Error", "Please add at least one valid source file with a range.")
            return False

        for i, frame in enumerate(valid_sources, start=1):
            if not os.path.isfile(frame.file_path):
                messagebox.showerror("Error", f"Source file {i} not found: {frame.file_path}")
                return False

        # Ensure each source can map to a track
        labels = self.get_track_labels()
        for i, sf in enumerate(self.source_frames, start=1):
            if sf.track_var.get() not in labels:
                messagebox.showerror("Error", f"Source {i} has an invalid Track selection.")
                return False

        return True

    # ----- Workspace Save/Load -----

    def save_workspace(self):
        if not self.validate_inputs():
            return
        save_path = filedialog.asksaveasfilename(
            title="Save Workspace",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not save_path:
            return

        data = {
            "base_file_path": self.base_file_path,
            "tracks": [bf.get_data() for bf in self.base_frames],
            "sources": [sf.get_data() for sf in self.source_frames],
        }
        try:
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self.log(f"Workspace saved: {os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save workspace: {e}")

    def load_workspace(self):
        load_path = filedialog.askopenfilename(
            title="Load Workspace",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not load_path:
            return

        try:
            with open(load_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load workspace: {e}")
            return

        # Clear current UI
        for f in self.base_frames:
            f.destroy()
        self.base_frames.clear()
        for f in self.source_frames:
            f.destroy()
        self.source_frames.clear()

        # Restore base file
        self.base_file_path = data.get("base_file_path", "")
        self.base_file_label.config(
            text=(os.path.basename(self.base_file_path) if self.base_file_path else "No file selected"),
            foreground=("black" if self.base_file_path else "gray")
        )

        # Restore tracks
        tracks = data.get("tracks", [])
        if not tracks:
            self.add_base_cell()
        else:
            for t in tracks:
                self.add_base_cell()
                self.base_frames[-1].set_data(t)

        # Restore sources
        sources = data.get("sources", [])
        if not sources:
            self.add_source_file()
        else:
            for s in sources:
                self.add_source_file()
                self.source_frames[-1].set_data(s)

            # Now apply track selections using labels (after all tracks exist)
            labels = self.get_track_labels()
            for sf, sdata in zip(self.source_frames, sources):
                lbl = sdata.get("track_label", "")
                if lbl in labels:
                    sf.track_var.set(lbl)
                else:
                    # default to first
                    sf.track_combo.current(0)

        self.root.update_idletasks()
        self.log(f"Workspace loaded: {os.path.basename(load_path)}")

    # ----- Processing -----

    def process_ranges(self):
        if not self.validate_inputs():
            return

        self.output_text.delete(1.0, tk.END)
        self.log("Starting range processing...")

        try:
            # Load base workbook
            self.log(f"[io] Loading base file: {os.path.basename(self.base_file_path)}")
            base_wb = load_workbook(self.base_file_path)

            # Build track info
            tracks_info = []  # list of dicts: { ws, base_col, base_row, current_col, fixed_row, start_ref, end_ref }
            for i, bf in enumerate(self.base_frames, start=1):
                t = bf.get_data()
                sheet_name, base_ref = parse_sheet_and_ref(t["base_cell"])
                base_ws = get_sheet(base_wb, sheet_name)
                base_col, base_row = parse_cell(base_ref)

                tracks_info.append({
                    "ws": base_ws,
                    "base_col": base_col,
                    "base_row": base_row,
                    "current_col": base_col,   # moving cursor
                    "fixed_row": base_row,     # row stays fixed for the track
                    "start_ref": t["start_ref"],  # optional "Sheet!Cell" or "Cell"
                    "end_ref": t["end_ref"],
                })
                self.log(f"[calc] Track #{i} base={t['base_cell']} (col={base_col}, row={base_row}) start_ref='{t['start_ref']}' end_ref='{t['end_ref']}'")

            # Load sources
            sources = []
            for i, sf in enumerate(self.source_frames, start=1):
                if sf.is_valid():
                    self.log(f"[io] Loading source file {i}: {os.path.basename(sf.file_path)}")
                    src_wb = load_workbook(sf.file_path, data_only=True)
                    sources.append({
                        "title": sf.title_entry.get().strip(),
                        "workbook": src_wb,
                        "range": sf.range_entry.get().strip(),
                        "filename": os.path.basename(sf.file_path),
                        "track_label": sf.track_var.get(),
                    })

            # Helper: read value from a reference in BASE workbook (to copy as search token)
            def read_value_from_base_ref(ref_str: str) -> Optional[Any]:
                if not ref_str:
                    log_debug(f"[calc] read_value_from_base_ref (empty) -> None")
                    return None
                sh, cellref = parse_sheet_and_ref(ref_str)
                ws = get_sheet(base_wb, sh)
                c, r = parse_cell(cellref)
                val = ws.cell(row=r, column=c).value
                log_debug(f"[calc] read_value_from_base_ref '{ref_str}' -> value='{normalize_value(val)}'")
                return val

            # Process each source in order, routing to its selected track
            for i, src in enumerate(sources, start=1):
                # Resolve which track this source follows
                labels = self.get_track_labels()
                try:
                    track_idx = labels.index(src["track_label"])
                except ValueError:
                    track_idx = 0  # fallback to first
                track = tracks_info[track_idx]

                base_ws = track["ws"]
                current_col = track["current_col"]
                fixed_row = track["fixed_row"]

                src_wb = src["workbook"]
                src_range = src["range"]
                filename = src["filename"]

                title = src.get("title") or ""
                nice_name = f"{title} ({filename})" if title else filename
                self.log(f"Processing source {i} {nice_name} on track #{track_idx+1} [{labels[track_idx]}]: {src_range}")

                # Parse source range into ws + boundaries
                src_sheet_name, src_ref = parse_sheet_and_ref(src_range)
                src_ws = get_sheet(src_wb, src_sheet_name)

                # Decide on wildcard-rows/cols resolution
                bounds: Tuple[int, int, int, int]
                if ":" in src_ref:
                    left, right = [x.strip() for x in src_ref.split(":", 1)]

                    # Parse tokens with wildcard support
                    left_col_tok, left_row_tok = split_col_row(left)
                    end_col_tok, end_row_tok = split_col_row(right)

                    # ---------- Resolve START COLUMN ----------
                    if left_col_tok is None:
                        raise ValueError(f'Invalid start column in "{left}"')
                    if left_col_tok == "?":
                        # If caller ever allows '?' start column, default to column A (1)
                        start_col = 1
                    else:
                        start_col = column_index_from_string(left_col_tok)

                    # Helper: reference-driven row lookup
                    def ref_row_or_none(ref_key: str) -> Optional[int]:
                        ref_str = track.get(ref_key)
                        if not ref_str:
                            return None
                        val = read_value_from_base_ref(ref_str)
                        if val is None or normalize_value(val) == "":
                            return None
                        return find_row_by_value(src_ws, val)

                    # ---------- Resolve START ROW ----------
                    if left_row_tok is None:
                        # No row component (e.g., just a column) — default to 1
                        start_row = 1
                    elif left_row_tok == "?":
                        # First try Start Row Ref
                        r = ref_row_or_none("start_ref")
                        if r is not None:
                            start_row = r
                            log_debug(f"[calc] start_row from start_ref -> {start_row}")
                        else:
                            # Fallback: first non-empty row on the start column
                            start_row = first_nonempty_row_on_col(src_ws, start_col, 1)
                            log_debug(f"[calc] start_row fallback first non-empty -> {start_row}")
                    else:
                        # Specific row number
                        start_row = int(left_row_tok)

                    # ---------- Resolve END COLUMN (initial pass) ----------
                    # If end_col is '?', we need a row span. We might not yet know end_row.
                    # Use a provisional max_row = ws.max_row; we'll refine after end_row is finalized.
                    if end_col_tok is None:
                        end_col = start_col
                    elif end_col_tok == "?":
                        provisional_min_r = start_row
                        provisional_max_r = src_ws.max_row
                        end_col = last_nonempty_col_in_row_range(src_ws, start_col, provisional_min_r,
                                                                 provisional_max_r)
                        log_debug(f"[calc] provisional end_col from '?' -> {end_col}")
                    else:
                        end_col = column_index_from_string(end_col_tok)

                    # ---------- Resolve END ROW ----------
                    if end_row_tok is None:
                        end_row = start_row
                    elif end_row_tok == "?":
                        # Try End Row Ref first
                        r_end = ref_row_or_none("end_ref")
                        if r_end is not None:
                            end_row = r_end
                            log_debug(f"[calc] end_row from end_ref -> {end_row}")
                        else:
                            # Need columns to compute last non-empty row across [start_col..end_col]
                            min_c_tmp, max_c_tmp = (start_col, end_col) if start_col <= end_col else (
                            end_col, start_col)
                            end_row = last_nonempty_row_in_col_range(src_ws, min_c_tmp, max_c_tmp, start_row)
                            log_debug(
                                f"[calc] end_row fallback last non-empty across cols {min_c_tmp}..{max_c_tmp} -> {end_row}")
                    else:
                        end_row = int(end_row_tok)

                    # ---------- If end_col was '?', refine it now using the resolved row range ----------
                    if end_col_tok == "?":
                        min_r_final = min(start_row, end_row)
                        max_r_final = max(start_row, end_row)
                        end_col = last_nonempty_col_in_row_range(src_ws, start_col, min_r_final, max_r_final)
                        log_debug(f"[calc] refined end_col with rows {min_r_final}..{max_r_final} -> {end_col}")

                    # Final bounds
                    min_c, max_c = sorted((start_col, end_col))
                    min_r, max_r = sorted((start_row, end_row))
                    bounds = (min_c, min_r, max_c, max_r)
                    self.log(f"[calc] resolved bounds -> rows {min_r}..{max_r}; cols {min_c}..{max_c}")
                    resolved_ref = format_range_a1(min_c, min_r, max_c, max_r, src_sheet_name)
                    self.log(f"[calc] resolved source range -> {resolved_ref}")
                else:
                    c, r = parse_cell(src_ref)
                    bounds = (c, r, c, r)
                    log_debug(f"[calc] single-cell source ref '{src_ref}' -> bounds=({c},{r})..({c},{r})")
                    resolved_ref = format_range_a1(c, r, c, r, src_sheet_name)
                    self.log(f"[calc] resolved source range -> {resolved_ref}")

                # Paste
                width, height = copy_values(
                    src_ws=src_ws,
                    dst_ws=base_ws,
                    src_bounds=bounds,
                    dst_row=fixed_row,
                    dst_col=current_col
                )

                # Log paste range
                start_letter = get_column_letter(current_col)
                end_letter = get_column_letter(current_col + width - 1)
                if height <= 1:
                    paste_range = f"{start_letter}{fixed_row}:{end_letter}{fixed_row}"
                else:
                    paste_range = f"{start_letter}{fixed_row}:{end_letter}{fixed_row + height - 1}"
                self.log(f"  → Pasted {width}x{height} to {paste_range}")

                # Advance that track's cursor
                track["current_col"] = current_col + width
                log_debug(f"[calc] advance track cursor from col={current_col} by width={width} -> col={track['current_col']}")

            # Save base workbook
            base_wb.save(self.base_file_path)
            self.log(f"[io] Saved changes to: {os.path.basename(self.base_file_path)}")
            self.log("✓ Processing completed successfully!")

        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Error", error_msg)

def main():
    root = tk.Tk()
    app = RangePasterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
