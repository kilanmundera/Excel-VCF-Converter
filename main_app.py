import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import threading
import sv_ttk

# Import the backend logic
try:
    from vcf_backend import create_vcf_from_excel, CONFIG
except ImportError:
    messagebox.showerror("Error", "backend.py not found.\nMake sure backend.py is in the same folder.")
    exit()

class VCFEditorApp:
    """
    Main application class for the Excel to VCF Converter.
    Handles all UI logic, state management, and background thread for conversion.
    """
    def __init__(self, root):
        """Initializes the application, sets up the theme, and builds the UI."""
        self.root = root
        self.root.title("Excel → VCF Converter")
        sv_ttk.set_theme("dark")

        # --- Internal State ---
        self.df = None
        self.excel_columns = ["(None)"]
        
        self.vcard_field_map = {
            "Phone": "Phone Number",
            "Full Name": "Full Name",
            "Title": "Title",
            "Email": "Email",
            "Notes": "Notes"
        }
        self.vcard_keys = list(self.vcard_field_map.keys()) # Cleaner definition
        
        # --- Widget Storage (Initialized to None) ---
        self.mapping_combos = {}
        self.template_buttons = {}
        
        # UI components
        self.file_status_label = None
        self.tree = None
        self.config_frame = None
        self.fn_template_entry = None
        self.clear_button = None
        self.convert_button = None

        # Build the main UI
        self._build_ui()

    def _build_ui(self):
        """Constructs the main UI layout using a grid system."""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True)
        # Configure grid for expansion in the main window
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        content_frame = ttk.Frame(main_frame, padding=10)
        content_frame.grid(row=0, column=0)
        
        # Construct sections
        self._build_file_loader(content_frame)
        self._build_preview(content_frame)
        self._build_mapping(content_frame)
        self._build_name_builder(content_frame)
        self._build_convert_button(content_frame)

    def _build_file_loader(self, parent):
        """Builds the UI for opening the Excel file."""
        frame = ttk.LabelFrame(parent, text="Import Excel File", padding=10)
        frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)
        
        center_frame = ttk.Frame(frame)
        center_frame.pack(anchor="center")
        
        open_button = ttk.Button(center_frame, text="Import", 
                                 command=self.open_excel, style="Accent.TButton")
        open_button.pack(pady=(0, 5)) 
        
        self.file_status_label = ttk.Label(center_frame, text="No file loaded.", 
                                           foreground="lightgray")
        self.file_status_label.pack()

    def _build_preview(self, parent):
        """Builds the 'Preview' table, fixing its size to prevent window expansion."""
        frame = ttk.LabelFrame(parent, text="File Preview (First 5 Rows)", padding=10)
        frame.grid(row=1, column=0, sticky="n", pady=5, padx=(0, 5))
        
        # Use pack_propagate to fix the frame size
        frame.pack_propagate(False) 
        frame.config(width=400, height=220) 
        
        self.tree = ttk.Treeview(frame, show="headings", height=6) 
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True) 
        self.tree.configure(xscrollcommand=hsb.set) # Correct way to link scrollbar

    def _build_mapping(self, parent):
        """Builds the column mapping section."""
        self.config_frame = ttk.LabelFrame(parent, text="Column Mapping", padding=10)
        self.config_frame.grid(row=1, column=1, sticky="nsew", pady=5, padx=(5, 0))
        
        # Initial placeholder label
        placeholder = ttk.Label(self.config_frame, text="......")
        placeholder.pack(anchor="center")
        
    def _build_name_builder(self, parent):
        """Builds the contact name builder with entry and helper buttons."""
        frame = ttk.LabelFrame(parent, text="Build Contact Name", padding=10)
        frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        
        ttk.Label(frame, text="Build the contact name with helper buttons.").pack(anchor="center")
        
        self.fn_template_entry = ttk.Entry(frame)
        self.fn_template_entry.pack(fill="x", pady=5, padx=5)
        
        button_frame = ttk.Frame(frame)
        button_frame.pack()
        
        # Build template helper buttons
        self.template_buttons = {}
        for key in self.vcard_keys:
            label = self.vcard_field_map[key]
            btn = ttk.Button(button_frame, text=label, 
                             command=lambda k=key: self.add_to_template(k),
                             state="disabled") 
            btn.pack(side="left", padx=2)
            self.template_buttons[key] = btn
            
        # Clear button (no Accent style for a secondary action)
        self.clear_button = ttk.Button(button_frame, text="Clear", 
                                       command=self.clear_template, 
                                       state="disabled")
        self.clear_button.pack(side="left", padx=(10, 2))
        
    def _build_convert_button(self, parent):
        """Builds the final 'Convert to VCF' button."""
        frame = ttk.Frame(parent, padding=(0, 10, 0, 0))
        frame.grid(row=3, column=0, columnspan=2, sticky="ew")
        
        self.convert_button = ttk.Button(frame, text="Convert", 
                                         command=self.start_conversion, 
                                         state="disabled", style="Accent.TButton")
        self.convert_button.pack(anchor="center") 

    # --- Core Application Logic ---

    def open_excel(self):
        """Opens a file dialog, loads the file, updates status, and populates UI."""
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path:
            return
            
        # Set loading status and force UI refresh
        self.file_status_label.config(text="Reading file, please wait...")
        self.root.update() 

        try:
            # Load with dtype=str to maintain exact formatting
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

    def show_dataframe(self, df):
        """Populates the Treeview with the first 5 rows of the DataFrame."""
        # Clear existing data
        self.tree.delete(*self.tree.get_children())
        
        # Set columns
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="w") 

        # Insert rows
        for _, row in df.head(5).iterrows():
            self.tree.insert("", "end", values=list(row))

    def populate_config_frame(self):
        """Dynamically creates the column mapping widgets after a file is loaded."""
        # Clear existing widgets in the config frame
        for widget in self.config_frame.winfo_children():
            widget.destroy()
        self.mapping_combos = {}

        for key in self.vcard_keys:
            label_text = self.vcard_field_map[key] 
            
            row_frame = ttk.Frame(self.config_frame)
            row_frame.pack(fill="x", pady=2)
            
            label = ttk.Label(row_frame, text=f"{label_text}:", width=15)
            label.pack(side="left")
            
            var = tk.StringVar(self.config_frame)
            var.set("(None)") 
            # Use self.check_mapping_validity as a direct callable
            var.trace_add("write", self.check_mapping_validity)
            
            combo = ttk.Combobox(row_frame, textvariable=var, 
                                 values=self.excel_columns, state="readonly")
            combo.pack(side="left", fill="x", expand=True, padx=5)
            
            self.mapping_combos[key] = var 
            
        self.check_mapping_validity()

    def check_mapping_validity(self, *args):
        """
        Checks all mappings and template.
        Enables/disables Convert, Clear, and individual name builder buttons.
        """
        any_mapping_valid = False
        
        # Determine which fields are mapped and update template buttons
        for key, var in self.mapping_combos.items():
            is_mapped = var.get() != "(None)"
            if is_mapped:
                any_mapping_valid = True
            
            if key in self.template_buttons:
                self.template_buttons[key].config(state="normal" if is_mapped else "disabled")
        
        # Determine global button state
        if any_mapping_valid and self.fn_template_entry.get().strip():
            self.convert_button.config(state="normal")
            self.clear_button.config(state="normal")
        else:
            self.convert_button.config(state="disabled")
            self.clear_button.config(state="disabled")

    def populate_structure_frame(self):
        """Sets the default name format after a file is loaded."""
        self.fn_template_entry.delete(0, tk.END)
        # Default template using the Full Name field
        self.fn_template_entry.insert(0, "{{Full Name}}")
        
        self.check_mapping_validity()

    def clear_template(self):
        """Clears the contact name template entry."""
        self.fn_template_entry.delete(0, tk.END)

    def add_to_template(self, field_key):
        """
        Appends a field placeholder (e.g., '{{Full Name}}') to the template entry.
        Adds a preceding space if the entry is not empty.
        """
        current_text = self.fn_template_entry.get()
        placeholder = f"{{{{{field_key}}}}}"
        
        # Check for empty string more strictly
        if not current_text.strip():
            self.fn_template_entry.insert(tk.END, placeholder)
        else:
            self.fn_template_entry.insert(tk.END, f" {placeholder}")
            
        # Re-check validity after adding text
        self.check_mapping_validity()


    def start_conversion(self):
        """
        Validates all user inputs, prepares the CONFIG, and starts the conversion
        process in a new, non-blocking thread.
        """
        # --- Pre-Conversion Validation ---
        if self.df is None:
            messagebox.showwarning("Warning", "Please open an Excel file first.")
            return

        column_mapping = {}
        for key, var in self.mapping_combos.items():
            selected_col = var.get()
            if selected_col != "(None)":
                column_mapping[key] = selected_col
        
        if not column_mapping:
            messagebox.showwarning("Warning", "Please map at least one column in Column Mapping.")
            return

        fn_template = self.fn_template_entry.get().strip()
        
        if not fn_template:
            messagebox.showwarning("Warning", "Please build a format in Contact Name Format.")
            return

        # --- File Save Dialog ---
        output_path = filedialog.asksaveasfilename(
            title="Save as VCF",
            defaultextension=".vcf",
            filetypes=[("VCF files", "*.vcf")],
            initialfile="Output.vcf"
        )
        if not output_path:
            return

        # --- Prepare CONFIG for backend ---
        CONFIG["output_vcf_filename"] = output_path
        CONFIG["dataframe_to_process"] = self.df
        CONFIG["column_mapping"] = column_mapping
        CONFIG["fn_template"] = fn_template

        # --- Start Thread ---
        self.file_status_label.config(text="Converting...")
        self.convert_button.config(state="disabled") # Disable button during conversion

        threading.Thread(target=self.run_conversion, daemon=True).start()

    def run_conversion(self):
        """
        Data-processing target for the background thread.
        Calls the backend function and handles success/error messages.
        Ensures UI updates (like re-enabling the button) happen on the main thread.
        """
        try:
            create_vcf_from_excel(CONFIG)
            
            # Use root.after for UI updates from the background thread
            self.root.after(0, lambda: self.file_status_label.config(text="✅ Conversion complete"))
            self.root.after(0, lambda: messagebox.showinfo("Success", "VCF file created successfully!"))
        except Exception as e:
            self.root.after(0, lambda: self.file_status_label.config(text="❌ Failed"))
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            # Re-enable the convert button and re-check validity on the main thread
            self.root.after(0, lambda: self.convert_button.config(state="normal"))
            self.root.after(0, self.check_mapping_validity)

# --- Run the Application ---
if __name__ == "__main__":
    root = tk.Tk()
    app = VCFEditorApp(root)
    root.mainloop()
