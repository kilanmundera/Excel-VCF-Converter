import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import threading
import sv_ttk

# Import the backend logic
try:
    from vcf_backend import create_vcf_from_excel, CONFIG, PHONE_KEYS
except ImportError:
    messagebox.showerror("Error", "vcf_backend.py not found.\nMake sure vcf_backend.py is in the same folder.")
    exit()

# Available phone types for the comboboxes
PHONE_TYPE_OPTIONS = ["CELL", "WORK", "HOME", "VOICE", "FAX", "PAGER"]

# Fixed widths
COL_COMBO_WIDTH  = 16
TYPE_COMBO_WIDTH = 6


class VCFEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel → VCF Converter")
        sv_ttk.set_theme("dark")

        # --- Internal State ---
        self.df = None
        self.excel_columns = ["(None)"]

        self.vcard_field_map = {
            "Full Name": "Full Name",
            "Title":     "Title",
            "Phone 1":   "Phone 1",
            "Phone 2":   "Phone 2",
            "Phone 3":   "Phone 3",
            "Email":     "Email",
            "Notes":     "Notes"
        }
        self.vcard_keys = list(self.vcard_field_map.keys())

        # --- Widget Storage ---
        self.mapping_combos   = {}
        self.phone_type_vars  = {}
        self.template_buttons = {}

        # Group widgets
        self.group_mode_var    = None   # "fixed" or "column"
        self.group_fixed_var   = None   # StringVar for the fixed name entry
        self.group_column_var  = None   # StringVar for the column combobox
        self.group_fixed_entry = None
        self.group_col_combo   = None

        self.file_status_label = None
        self.tree              = None
        self.config_frame      = None
        self.fn_template_entry = None
        self.clear_button      = None
        self.convert_button    = None

        self._build_ui()

    # ------------------------------------------------------------------ #
    #  UI CONSTRUCTION                                                     #
    # ------------------------------------------------------------------ #

    def _build_ui(self):
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        content = ttk.Frame(self.root, padding=10)
        content.grid(row=0, column=0, sticky="nsew")

        content.rowconfigure(0, weight=0)
        content.rowconfigure(1, weight=1)
        content.rowconfigure(2, weight=0)
        content.rowconfigure(3, weight=0)
        content.rowconfigure(4, weight=0)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=0)

        self._build_file_loader(content)
        self._build_preview(content)
        self._build_mapping(content)
        self._build_name_builder(content)
        self._build_group_section(content)
        self._build_convert_button(content)

    def _build_file_loader(self, parent):
        frame = ttk.LabelFrame(parent, text="Import Excel File", padding=10)
        frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))

        center = ttk.Frame(frame)
        center.pack(anchor="center")

        ttk.Button(center, text="Import",
                   command=self.open_excel,
                   style="Accent.TButton").pack(pady=(0, 5))

        self.file_status_label = ttk.Label(center, text="No file loaded.",
                                           foreground="lightgray")
        self.file_status_label.pack()

    def _build_preview(self, parent):
        frame = ttk.LabelFrame(parent, text="File Preview (First 5 Rows)", padding=10)
        frame.grid(row=1, column=0, sticky="nsew", pady=5, padx=(0, 5))

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, show="headings")
        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def _build_mapping(self, parent):
        self.config_frame = ttk.LabelFrame(parent, text="Column Mapping", padding=10)
        self.config_frame.grid(row=1, column=1, sticky="ns", pady=5, padx=(5, 0))

        ttk.Label(self.config_frame, text="......").pack(anchor="center")

    def _build_name_builder(self, parent):
        frame = ttk.LabelFrame(parent, text="Build Contact Name", padding=10)
        frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Build the contact name with helper buttons.").grid(
            row=0, column=0, sticky="w")

        self.fn_template_entry = ttk.Entry(frame)
        self.fn_template_entry.grid(row=1, column=0, sticky="ew", pady=5, padx=5)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0)

        phone_keys = set(PHONE_KEYS)
        self.template_buttons = {}
        for key in self.vcard_keys:
            if key in phone_keys:
                continue
            btn = ttk.Button(btn_frame, text=self.vcard_field_map[key],
                             command=lambda k=key: self.add_to_template(k),
                             state="disabled")
            btn.pack(side="left", padx=2)
            self.template_buttons[key] = btn

        self.clear_button = ttk.Button(btn_frame, text="Clear",
                                       command=self.clear_template,
                                       state="disabled")
        self.clear_button.pack(side="left", padx=(10, 2))

    def _build_group_section(self, parent):
        """
        Group / CATEGORIES section.
        Two modes:
          - Fixed  : all contacts get the same group name (typed in an Entry)
          - Column : group name is read from an Excel column (Combobox)
        A third radio "None" disables grouping entirely.
        """
        frame = ttk.LabelFrame(parent, text="Contact Group (CATEGORIES)", padding=10)
        frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5)
        frame.columnconfigure(1, weight=1)

        self.group_mode_var   = tk.StringVar(value="none")
        self.group_fixed_var  = tk.StringVar()
        self.group_column_var = tk.StringVar(value="(None)")

        # --- Radio: None ---
        ttk.Radiobutton(frame, text="No group",
                        variable=self.group_mode_var, value="none",
                        command=self._on_group_mode_change).grid(
            row=0, column=0, sticky="w", padx=(0, 15))

        # --- Radio: Fixed name ---
        ttk.Radiobutton(frame, text="Fixed name:",
                        variable=self.group_mode_var, value="fixed",
                        command=self._on_group_mode_change).grid(
            row=0, column=1, sticky="w")

        self.group_fixed_entry = ttk.Entry(frame, textvariable=self.group_fixed_var,
                                           width=20, state="disabled")
        self.group_fixed_entry.grid(row=0, column=2, sticky="w", padx=(5, 20))

        # --- Radio: From column ---
        ttk.Radiobutton(frame, text="From column:",
                        variable=self.group_mode_var, value="column",
                        command=self._on_group_mode_change).grid(
            row=0, column=3, sticky="w")

        self.group_col_combo = ttk.Combobox(frame, textvariable=self.group_column_var,
                                            values=["(None)"],
                                            state="disabled", width=16)
        self.group_col_combo.grid(row=0, column=4, sticky="w", padx=(5, 0))

    def _on_group_mode_change(self):
        mode = self.group_mode_var.get()
        self.group_fixed_entry.config(state="normal"   if mode == "fixed"  else "disabled")
        self.group_col_combo.config( state="readonly"  if mode == "column" else "disabled")

    def _build_convert_button(self, parent):
        frame = ttk.Frame(parent, padding=(0, 10, 0, 0))
        frame.grid(row=4, column=0, columnspan=2, sticky="ew")

        self.convert_button = ttk.Button(frame, text="Convert",
                                         command=self.start_conversion,
                                         state="disabled",
                                         style="Accent.TButton")
        self.convert_button.pack(anchor="center")

    # ------------------------------------------------------------------ #
    #  CORE APPLICATION LOGIC                                              #
    # ------------------------------------------------------------------ #

    def open_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path:
            return

        self.file_status_label.config(text="Reading file, please wait...")
        self.root.update()

        try:
            df = pd.read_excel(path, dtype=str)
            df.columns = df.columns.str.strip()
            df = df.fillna("")
        except Exception as e:
            self.file_status_label.config(text="No file loaded.")
            messagebox.showerror("Error", f"Could not read file:\n{e}")
            return

        self.df = df
        self.excel_columns = ["(None)"] + list(df.columns)

        self.file_status_label.config(text=f"Loaded: {os.path.basename(path)}")
        self.show_dataframe(df)
        self.populate_config_frame()
        self.populate_structure_frame()

        # Update the group column combobox with the new columns
        self.group_col_combo.config(values=self.excel_columns)
        self.group_column_var.set("(None)")

    def show_dataframe(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="w")
        for _, row in df.head(5).iterrows():
            self.tree.insert("", "end", values=list(row))

    def populate_config_frame(self):
        for widget in self.config_frame.winfo_children():
            widget.destroy()
        self.mapping_combos  = {}
        self.phone_type_vars = {}

        for key in self.vcard_keys:
            label_text = self.vcard_field_map[key]
            is_phone   = key in PHONE_KEYS

            row_frame = ttk.Frame(self.config_frame)
            row_frame.pack(anchor="w", pady=2)

            ttk.Label(row_frame, text=f"{label_text}:", width=10).pack(side="left")

            col_var = tk.StringVar(value="(None)")
            col_var.trace_add("write", self.check_mapping_validity)
            ttk.Combobox(row_frame, textvariable=col_var,
                         values=self.excel_columns,
                         state="readonly",
                         width=COL_COMBO_WIDTH).pack(side="left", padx=(5, 0))
            self.mapping_combos[key] = col_var

            if is_phone:
                type_var = tk.StringVar(value="CELL")
                ttk.Combobox(row_frame, textvariable=type_var,
                             values=PHONE_TYPE_OPTIONS,
                             state="readonly",
                             width=TYPE_COMBO_WIDTH).pack(side="left", padx=(6, 0))
                ttk.Label(row_frame, text="type",
                          foreground="gray").pack(side="left", padx=(3, 0))
                self.phone_type_vars[key] = type_var

        self.check_mapping_validity()

    def check_mapping_validity(self, *args):
        any_mapped = False
        for key, var in self.mapping_combos.items():
            is_mapped = var.get() != "(None)"
            if is_mapped:
                any_mapped = True
            if key in self.template_buttons:
                self.template_buttons[key].config(
                    state="normal" if is_mapped else "disabled")

        ok = any_mapped and bool(self.fn_template_entry.get().strip())
        self.convert_button.config(state="normal" if ok else "disabled")
        self.clear_button.config(state="normal"   if ok else "disabled")

    def populate_structure_frame(self):
        self.fn_template_entry.delete(0, tk.END)
        self.fn_template_entry.insert(0, "{{Full Name}}")
        self.check_mapping_validity()

    def clear_template(self):
        self.fn_template_entry.delete(0, tk.END)

    def add_to_template(self, field_key):
        current = self.fn_template_entry.get()
        placeholder = f"{{{{{field_key}}}}}"
        if not current.strip():
            self.fn_template_entry.insert(tk.END, placeholder)
        else:
            self.fn_template_entry.insert(tk.END, f" {placeholder}")
        self.check_mapping_validity()

    def _build_group_config(self):
        """Builds the group_config dict to pass to the backend."""
        mode = self.group_mode_var.get()
        if mode == "none":
            return None
        if mode == "fixed":
            val = self.group_fixed_var.get().strip()
            if not val:
                messagebox.showwarning("Warning",
                                       "Please enter a group name, or select 'No group'.")
                return False   # signal validation failure
            return {"mode": "fixed", "value": val}
        if mode == "column":
            col = self.group_column_var.get()
            if col == "(None)":
                messagebox.showwarning("Warning",
                                       "Please select a column for the group, or select 'No group'.")
                return False
            return {"mode": "column", "value": col}
        return None

    def start_conversion(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please open an Excel file first.")
            return

        column_mapping = {k: v.get() for k, v in self.mapping_combos.items()
                          if v.get() != "(None)"}
        if not column_mapping:
            messagebox.showwarning("Warning",
                                   "Please map at least one column in Column Mapping.")
            return

        fn_template = self.fn_template_entry.get().strip()
        if not fn_template:
            messagebox.showwarning("Warning",
                                   "Please build a format in Contact Name Format.")
            return

        group_config = self._build_group_config()
        if group_config is False:   # validation failed inside _build_group_config
            return

        phone_types = {key: var.get() for key, var in self.phone_type_vars.items()}

        output_path = filedialog.asksaveasfilename(
            title="Save as VCF",
            defaultextension=".vcf",
            filetypes=[("VCF files", "*.vcf")],
            initialfile="Output.vcf"
        )
        if not output_path:
            return

        CONFIG["output_vcf_filename"]  = output_path
        CONFIG["dataframe_to_process"] = self.df
        CONFIG["column_mapping"]       = column_mapping
        CONFIG["fn_template"]          = fn_template
        CONFIG["phone_types"]          = phone_types
        CONFIG["group_config"]         = group_config

        self.file_status_label.config(text="Converting...")
        self.convert_button.config(state="disabled")
        threading.Thread(target=self.run_conversion, daemon=True).start()

    def run_conversion(self):
        try:
            create_vcf_from_excel(CONFIG)
            self.root.after(0, lambda: self.file_status_label.config(
                text="✅ Conversion complete"))
            self.root.after(0, lambda: messagebox.showinfo(
                "Success", "VCF file created successfully!"))
        except Exception as e:
            self.root.after(0, lambda: self.file_status_label.config(text="❌ Failed"))
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: self.convert_button.config(state="normal"))
            self.root.after(0, self.check_mapping_validity)


# --- Run the Application ---
if __name__ == "__main__":
    root = tk.Tk()
    app = VCFEditorApp(root)
    root.mainloop()
