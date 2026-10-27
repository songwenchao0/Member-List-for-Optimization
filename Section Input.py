import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


class SectionSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AISC Section Selector")
        self.root.minsize(760, 500)

        self.shapes = None
        self.shape_column = "AISC_Manual_Label"
        self.properties = [
            "AISC_Manual_Label",
            "W", "A", "d", "bf", "tw", "tf",
            "Ix", "Sx", "Zx", "Iy", "Sy", "Zy", "J",
        ]
        self.available_properties = []
        self.selected_sections = set()
        self.shape_names = None

        self.workbook_label = ttk.Label(root, text="No workbook selected")
        self.workbook_label.pack(fill="x", padx=12, pady=(12, 4))

        self.load_button = ttk.Button(
            root, text="Choose AISC workbook...", command=self.load_workbook
        )
        self.load_button.pack(anchor="w", padx=12, pady=(0, 12))

        range_frame = ttk.LabelFrame(root, text="Add W-shapes by d range")
        range_frame.pack(fill="x", padx=12, pady=4)

        ttk.Label(range_frame, text="Minimum d:").grid(
            row=0, column=0, padx=(10, 4), pady=10
        )
        self.min_depth_entry = ttk.Entry(range_frame, width=12)
        self.min_depth_entry.grid(row=0, column=1, padx=4, pady=10)

        ttk.Label(range_frame, text="Maximum d:").grid(
            row=0, column=2, padx=(14, 4), pady=10
        )
        self.max_depth_entry = ttk.Entry(range_frame, width=12)
        self.max_depth_entry.grid(row=0, column=3, padx=4, pady=10)

        self.range_button = ttk.Button(
            range_frame,
            text="Add matching W-shapes",
            command=self.add_w_shapes_by_depth,
            state="disabled",
        )
        self.range_button.grid(row=0, column=4, padx=10, pady=10)

        specific_frame = ttk.LabelFrame(root, text="Add sections")
        specific_frame.pack(fill="x", padx=12, pady=4)

        self.shape_entry = ttk.Entry(specific_frame)
        self.shape_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        self.shape_entry.bind("<Return>", lambda _event: self.add_specific_shape())

        self.add_button = ttk.Button(
            specific_frame,
            text="Add section",
            command=self.add_specific_shape,
            state="disabled",
        )
        self.add_button.pack(side="left", padx=(0, 10), pady=10)

        self.add_matching_button = ttk.Button(
            specific_frame,
            text="Add all matching",
            command=self.add_matching_shapes,
            state="disabled",
        )
        self.add_matching_button.pack(side="left", padx=(0, 10), pady=10)

        selected_frame = ttk.LabelFrame(root, text="Selected sections")
        selected_frame.pack(fill="both", expand=True, padx=12, pady=4)

        self.table = ttk.Treeview(
            selected_frame, columns=(), show="headings", height=10
        )

        vertical_scrollbar = ttk.Scrollbar(
            selected_frame, orient="vertical", command=self.table.yview
        )
        horizontal_scrollbar = ttk.Scrollbar(
            selected_frame, orient="horizontal", command=self.table.xview
        )
        self.table.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set,
        )
        self.table.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        selected_frame.rowconfigure(0, weight=1)
        selected_frame.columnconfigure(0, weight=1)

        action_frame = ttk.Frame(root)
        action_frame.pack(fill="x", padx=12, pady=(4, 8))

        self.remove_button = ttk.Button(
            action_frame,
            text="Remove selected",
            command=self.remove_selected,
            state="disabled",
        )
        self.remove_button.pack(side="left")

        self.export_button = ttk.Button(
            action_frame,
            text="Export CSV...",
            command=self.export_csv,
            state="disabled",
        )
        self.export_button.pack(side="right")

        self.status_label = ttk.Label(root, text="Choose an AISC workbook to begin.")
        self.status_label.pack(fill="x", padx=12, pady=(0, 10))

    def load_workbook(self):
        workbook = filedialog.askopenfilename(
            title="Select the AISC Shapes Database Excel file",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
        if not workbook:
            return

        try:
            with pd.ExcelFile(workbook) as excel:
                shapes = None
                for sheet_name in excel.sheet_names:
                    candidate = pd.read_excel(excel, sheet_name=sheet_name)
                    if self.shape_column in candidate.columns:
                        shapes = candidate
                        break

                if shapes is None:
                    raise ValueError(
                        f"Couldn't find '{self.shape_column}' in any worksheet. "
                        f"Available worksheets: {', '.join(excel.sheet_names)}"
                    )
        except (OSError, ValueError, ImportError) as error:
            messagebox.showerror("Could not load workbook", str(error))
            return

        self.shapes = shapes
        self.available_properties = [
            column for column in self.properties if column in shapes.columns
        ]
        self.shape_names = shapes[self.shape_column].astype(str).str.strip().str.upper()
        self.table.configure(columns=self.available_properties)
        for column in self.available_properties:
            self.table.heading(column, text=column)
            self.table.column(column, width=85, minwidth=65, stretch=True)
        self.workbook_label.config(text=workbook)
        self.range_button.config(state="normal")
        self.add_button.config(state="normal")
        self.add_matching_button.config(state="normal")
        self.status_label.config(text="Workbook loaded. Add sections to the list.")

    def add_sections(self, sections):
        added_count = 0
        for section in sections:
            if not section or section in self.selected_sections:
                continue
            match = self.shapes.loc[self.shape_names == section]
            if match.empty:
                continue

            self.selected_sections.add(section)
            row = match.iloc[0]
            values = [row[column] for column in self.available_properties]
            self.table.insert("", "end", iid=section, values=values)
            added_count += 1

        self.remove_button.config(
            state="normal" if self.selected_sections else "disabled"
        )
        self.export_button.config(
            state="normal" if self.selected_sections else "disabled"
        )
        return added_count

    def add_w_shapes_by_depth(self):
        if self.shapes is None:
            return

        min_value = self.min_depth_entry.get().strip()
        max_value = self.max_depth_entry.get().strip()
        if not min_value and not max_value:
            self.status_label.config(
                text="No d range selected. Add specific sections below."
            )
            return

        try:
            min_depth = float(min_value)
            max_depth = float(max_value)
        except ValueError:
            messagebox.showerror(
                "Invalid d range",
                "Enter both a minimum and maximum d value, or leave both blank "
                "and add specific sections instead.",
            )
            return

        if min_depth > max_depth:
            messagebox.showerror(
                "Invalid d range", "Minimum d cannot be greater than maximum d."
            )
            return

        required_columns = {"Type", "d"}
        missing_columns = required_columns.difference(self.shapes.columns)
        if missing_columns:
            messagebox.showerror(
                "Missing workbook data",
                "The selected worksheet is missing required columns: "
                f"{', '.join(sorted(missing_columns))}",
            )
            return

        depths = pd.to_numeric(self.shapes["d"], errors="coerce")
        is_w_shape = (
            self.shapes["Type"].astype(str).str.strip().str.upper() == "W"
        )
        matches = self.shapes.loc[
            is_w_shape & (depths >= min_depth) & (depths <= max_depth),
            self.shape_column,
        ]
        sections = matches.astype(str).str.strip().str.upper()
        added_count = self.add_sections(sections)
        self.status_label.config(
            text=f"Added {added_count} W-shapes with d from {min_depth:g} "
            f"to {max_depth:g}."
        )

    def add_specific_shape(self):
        if self.shapes is None:
            return

        section = self.shape_entry.get().strip().upper()
        if not section:
            messagebox.showinfo("Enter a section", "Type a section name to add it.")
            return
        if not (self.shape_names == section).any():
            messagebox.showerror(
                "Section not found", f"No AISC section named {section} was found."
            )
            return
        if section in self.selected_sections:
            self.status_label.config(text=f"{section} is already in the list.")
            return

        self.add_sections([section])
        self.shape_entry.delete(0, "end")
        self.status_label.config(text=f"Added {section}.")

    def add_matching_shapes(self):
        if self.shapes is None:
            return

        prefix = self.shape_entry.get().strip().upper()
        if not prefix:
            messagebox.showinfo(
                "Enter a shape prefix",
                "Enter a prefix such as W24 to add all matching sections.",
            )
            return

        matching_sections = self.shape_names.loc[
            self.shape_names.str.startswith(prefix, na=False)
        ]
        if matching_sections.empty:
            self.status_label.config(
                text=f"No sections found starting with {prefix}."
            )
            return

        added_count = self.add_sections(matching_sections)
        self.shape_entry.delete(0, "end")
        self.status_label.config(
            text=f"Added {added_count} section(s) starting with {prefix}."
        )

    def remove_selected(self):
        for item in self.table.selection():
            self.table.delete(item)
            self.selected_sections.discard(item)

        self.remove_button.config(
            state="normal" if self.selected_sections else "disabled"
        )
        self.export_button.config(
            state="normal" if self.selected_sections else "disabled"
        )
        self.status_label.config(
            text=f"{len(self.selected_sections)} sections selected."
        )

    def export_csv(self):
        if not self.selected_sections:
            return

        destination = filedialog.asksaveasfilename(
            title="Export selected sections",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="selected_sections.csv",
        )
        if not destination:
            return

        selected_rows = self.shapes.loc[
            self.shape_names.isin(self.selected_sections), self.available_properties
        ]
        try:
            selected_rows.to_csv(destination, index=False)
        except OSError as error:
            messagebox.showerror("Could not export CSV", str(error))
            return

        self.status_label.config(text=f"Exported {len(selected_rows)} sections.")
        messagebox.showinfo("Export complete", f"Selected sections saved to:\n{destination}")


root = tk.Tk()
app = SectionSelectorApp(root)
root.mainloop()
