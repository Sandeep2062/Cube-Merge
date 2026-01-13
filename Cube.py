import openpyxl
from openpyxl.drawing.image import Image as XLImage
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import winsound
from copy import deepcopy
import shutil
import sys
import webbrowser
import winreg

"""
╔══════════════════════════════════════════════════════════════════╗
║                    CUBE DATA PROCESSOR                            ║
║                                                                   ║
║  Developer: Sandeep (https://github.com/Sandeep2062)            ║
║  Repository: https://github.com/Sandeep2062/Cube-Data-Processor ║
║                                                                   ║
║  © 2026 Sandeep - All Rights Reserved                           ║
╚══════════════════════════════════════════════════════════════════╝
"""

# Set appearance mode and color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Resource path function for PyInstaller compatibility
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Registry-based Settings Management
class RegistrySettings:
    def __init__(self):
        self.SOFTWARE_KEY = r"SOFTWARE\CubeDataProcessor"
        self.app_key = None
        
    def _open_key(self, write=False):
        """Open or create registry key"""
        try:
            if write:
                self.app_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.SOFTWARE_KEY)
            else:
                self.app_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.SOFTWARE_KEY, 0, winreg.KEY_READ)
            return True
        except WindowsError:
            try:
                # Try to create the key if it doesn't exist
                self.app_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.SOFTWARE_KEY)
                return True
            except:
                return False
    
    def _close_key(self):
        """Close registry key"""
        if self.app_key:
            winreg.CloseKey(self.app_key)
            self.app_key = None
    
    def save_setting(self, name, value):
        """Save a setting to registry"""
        if self._open_key(write=True):
            try:
                if isinstance(value, list):
                    # Convert list to string with | separator
                    value = "|".join(value)
                winreg.SetValueEx(self.app_key, name, 0, winreg.REG_SZ, str(value))
            except:
                pass
            finally:
                self._close_key()
    
    def load_setting(self, name, default=""):
        """Load a setting from registry"""
        if self._open_key():
            try:
                value, _ = winreg.QueryValueEx(self.app_key, name)
                if "|" in str(value) and name == "grade_files":
                    # Convert string back to list
                    return str(value).split("|") if value else []
                return value
            except WindowsError:
                return default
            finally:
                self._close_key()
        return default
    
    def save_all_settings(self, grade_files, output_path, calendar_path):
        """Save all settings at once"""
        self.save_setting("grade_files", grade_files)
        self.save_setting("output_path", output_path)
        self.save_setting("calendar_path", calendar_path)
    
    def load_all_settings(self):
        """Load all settings at once"""
        grade_files = self.load_setting("grade_files", [])
        output_path = self.load_setting("output_path", "")
        calendar_path = self.load_setting("calendar_path", "")
        return {"output_path": output_path, "calendar_path": calendar_path}, grade_files

# Create global settings instance
registry_settings = RegistrySettings()

# Smart grade extraction
def extract_grade(filename):
    name = os.path.basename(filename).split('.')[0].upper()
    
    if "MORTAR" in name and "_" in name:
        parts = name.split("_")
        if len(parts) >= 3:
            ratio = f"{parts[-2]}:{parts[-1]}"
            return ratio
    
    name = name.replace("_", "").replace("-", "")
    return name.strip()

# FILLED ROW CHECKER
def get_last_row(ws):
    row = 2
    while True:
        if ws.cell(row=row, column=2).value in (None, ""):
            return row - 1
        row += 1

# SAFE WORKBOOK LOADING
def load_workbook_safe(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, keep_vba=False, data_only=False, keep_links=False)
        return wb
    except:
        wb = openpyxl.load_workbook(filepath)
        return wb

# LOAD CALENDAR DATA
def load_calendar_data(calendar_file, log_callback):
    try:
        if not calendar_file or not os.path.exists(calendar_file):
            log_callback("⚠ No calendar file selected")
            return None
        
        wb = load_workbook_safe(calendar_file)
        ws = wb.active
        
        calendar_dict = {}
        row = 2
        
        while True:
            casting_date = ws.cell(row=row, column=1).value
            if not casting_date:
                break
            
            date_7 = ws.cell(row=row, column=2).value
            date_28 = ws.cell(row=row, column=3).value
            
            if casting_date:
                date_str = str(casting_date).strip()
                calendar_dict[date_str] = {
                    "7_days": str(date_7).strip() if date_7 else "",
                    "28_days": str(date_28).strip() if date_28 else ""
                }
            
            row += 1
        
        wb.close()
        log_callback(f"✓ Calendar loaded: {len(calendar_dict)} dates")
        return calendar_dict
        
    except Exception as e:
        log_callback(f"✖ Calendar load error: {e}")
        return None

# PROCESS WITH GRADE AND/OR DATE
def process_combined(grade_files, office_file, output_folder, calendar_file, mode, log_callback):
    try:
        log_callback(f"\n{'='*60}")
        log_callback(f"PROCESSING MODE: {mode.upper().replace('_', ' ')}")
        log_callback(f"{'='*60}")
        
        calendar_data = None
        if mode in ["date_only", "both"]:
            calendar_data = load_calendar_data(calendar_file, log_callback)
            if not calendar_data:
                log_callback("✖ Cannot proceed without calendar file")
                return 0
        
        base = os.path.basename(office_file).split(".")[0]
        outname = f"{base}_Processed.xlsx"
        outpath = os.path.join(output_folder, outname)
        
        shutil.copy2(office_file, outpath)
        office_wb = load_workbook_safe(outpath)
        
        total_copy_count = 0
        
        # GRADE PROCESSING
        if mode in ["grade_only", "both"] and grade_files:
            log_callback(f"\n--- GRADE PROCESSING ---")
            
            for grade_file in grade_files:
                grade_wb = load_workbook_safe(grade_file)
                grade_ws = grade_wb.active
                grade_name = extract_grade(grade_file)
                
                log_callback(f"\nProcessing: {os.path.basename(grade_file)}")
                log_callback(f"Looking for grade: {grade_name}")
                
                last_row = get_last_row(grade_ws)
                log_callback(f"Data rows: {last_row - 1}")
                
                matching_sheets = []
                for sheet_name in office_wb.sheetnames:
                    ws = office_wb[sheet_name]
                    b12_value = ws["B12"].value
                    if b12_value:
                        b12 = str(b12_value).replace(" ", "").upper()
                        grade_normalized = grade_name.replace(" ", "").upper()
                        if b12 == grade_normalized:
                            matching_sheets.append(sheet_name)
                            log_callback(f"  ✓ Matched sheet: {sheet_name} (B12={b12})")
                
                log_callback(f"Total matching sheets: {len(matching_sheets)}")
                
                if len(matching_sheets) == 0:
                    log_callback(f"⚠ No sheets found with B12='{grade_name}'")
                    grade_wb.close()
                    continue
                
                sheet_index = 0
                
                for r in range(2, last_row + 1):
                    if sheet_index >= len(matching_sheets):
                        log_callback(f"⚠ More data rows than available sheets")
                        break
                    
                    current_sheet_name = matching_sheets[sheet_index]
                    ws = office_wb[current_sheet_name]
                    
                    weight_values = [grade_ws.cell(row=r, column=c).value for c in range(2, 8)]
                    strength_values = [grade_ws.cell(row=r, column=c).value for c in range(9, 15)]
                    
                    for i, v in enumerate(weight_values):
                        ws.cell(row=25, column=3 + i, value=v)
                    for i, v in enumerate(strength_values):
                        ws.cell(row=27, column=3 + i, value=v)
                    
                    total_copy_count += 1
                    log_callback(f"  ✓ Row {r} → {current_sheet_name}")
                    sheet_index += 1
                
                grade_wb.close()
        
        # DATE PROCESSING
        if mode in ["date_only", "both"] and calendar_data:
            log_callback(f"\n--- DATE PROCESSING ---")
            
            updated_count = 0
            
            for sheet_name in office_wb.sheetnames:
                ws = office_wb[sheet_name]
                
                casting_date_cell = ws["C17"].value
                if not casting_date_cell:
                    continue
                
                casting_date = str(casting_date_cell).strip()
                
                if casting_date in calendar_data:
                    date_7 = calendar_data[casting_date]["7_days"]
                    date_28 = calendar_data[casting_date]["28_days"]
                    
                    if date_7:
                        ws["C18"] = date_7
                    if date_28:
                        ws["F18"] = date_28
                    
                    updated_count += 1
                    log_callback(f"✓ {sheet_name}: {casting_date} → 7d:{date_7}, 28d:{date_28}")
                else:
                    log_callback(f"⚠ Date not in calendar: {casting_date} ({sheet_name})")
            
            log_callback(f"\nSheets updated: {updated_count}")
        
        office_wb.save(outpath)
        office_wb.close()
        
        log_callback(f"\n{'='*60}")
        log_callback(f"✓✓✓ SAVED: {outpath}")
        log_callback(f"{'='*60}")
        
        return total_copy_count
        
    except Exception as e:
        log_callback(f"✖ ERROR: {e}")
        import traceback
        log_callback(traceback.format_exc())
        return 0

class CubeDataProcessor:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Cube Data Processor")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Try to set icon using resource_path
        try:
            self.root.iconbitmap(resource_path("icon.ico"))
        except:
            pass
        
        # Load settings from registry
        settings, saved_grade_files = registry_settings.load_all_settings()
        
        # Variables
        self.grade_files = []
        self.office_path = ctk.StringVar()
        self.output_path = ctk.StringVar(value=settings.get("output_path", ""))
        self.calendar_path = ctk.StringVar(value=settings.get("calendar_path", ""))
        self.mode_var = ctk.StringVar(value="both")
        
        # Load saved grade files
        for gf in saved_grade_files:
            if os.path.exists(gf):
                self.grade_files.append(gf)
        
        self.setup_ui()
        
    def setup_ui(self):
        # Configure grid weights
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
        # Create sidebar
        self.create_sidebar()
        
        # Create main content area
        self.create_main_content()
        
        # Create footer
        self.create_footer()
        
    def create_sidebar(self):
        # Sidebar frame
        self.sidebar = ctk.CTkFrame(self.root, width=280, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar.grid_rowconfigure(8, weight=1)
        
        # Logo/Title in sidebar
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Logo using resource_path
        try:
            from PIL import Image, ImageTk
            logo_img = Image.open(resource_path("logo.png"))
            logo_img = logo_img.resize((60, 60), Image.Resampling.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_img)
            logo_label = ctk.CTkLabel(logo_frame, image=logo_photo, text="")
            logo_label.image = logo_photo
            logo_label.pack()
        except:
            logo_label = ctk.CTkLabel(logo_frame, text="🔷", font=ctk.CTkFont(size=40))
            logo_label.pack()
        
        # Title
        title_label = ctk.CTkLabel(logo_frame, text="CUBE DATA\nPROCESSOR", 
                                 font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(10, 0))
        
        # Mode selection
        mode_label = ctk.CTkLabel(self.sidebar, text="Processing Mode", 
                                font=ctk.CTkFont(size=14, weight="bold"))
        mode_label.grid(row=1, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.grade_radio = ctk.CTkRadioButton(self.sidebar, text="📊 Grade Only", 
                                             variable=self.mode_var, value="grade_only",
                                             command=self.update_mode_ui)
        self.grade_radio.grid(row=2, column=0, padx=20, pady=5, sticky="w")
        
        self.date_radio = ctk.CTkRadioButton(self.sidebar, text="📅 Date Only", 
                                            variable=self.mode_var, value="date_only",
                                            command=self.update_mode_ui)
        self.date_radio.grid(row=3, column=0, padx=20, pady=5, sticky="w")
        
        self.both_radio = ctk.CTkRadioButton(self.sidebar, text="🔄 Both", 
                                            variable=self.mode_var, value="both",
                                            command=self.update_mode_ui)
        self.both_radio.grid(row=4, column=0, padx=20, pady=5, sticky="w")
        
        # Grade files section
        grade_label = ctk.CTkLabel(self.sidebar, text="Grade Files", 
                                  font=ctk.CTkFont(size=14, weight="bold"))
        grade_label.grid(row=5, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.grade_listbox = ctk.CTkTextbox(self.sidebar, height=100)
        self.grade_listbox.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        # Update grade listbox
        self.update_grade_listbox()
        
        # Grade file buttons
        grade_btn_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        grade_btn_frame.grid(row=7, column=0, padx=20, pady=(0, 20), sticky="ew")
        
        add_grade_btn = ctk.CTkButton(grade_btn_frame, text="➕ Add", 
                                      command=self.add_grades, width=100)
        add_grade_btn.pack(side="left", padx=(0, 5))
        
        clear_grade_btn = ctk.CTkButton(grade_btn_frame, text="🗑️ Clear", 
                                        command=self.clear_grades, width=100)
        clear_grade_btn.pack(side="left")
        
        # Social links
        social_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        social_frame.grid(row=9, column=0, padx=20, pady=20)
        
        github_btn = ctk.CTkButton(social_frame, text="🐙 GitHub", 
                                  command=lambda: webbrowser.open("https://github.com/Sandeep2062/Cube-Data-Processor"),
                                  width=110)
        github_btn.pack(side="left", padx=5)
        
        insta_btn = ctk.CTkButton(social_frame, text="📷 Instagram", 
                                 command=lambda: webbrowser.open("https://www.instagram.com/sandeep._.2062/"),
                                 width=110, fg_color="#E1306C", hover_color="#C13584")
        insta_btn.pack(side="left", padx=5)
        
    def create_main_content(self):
        # Main content frame
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.grid(row=1, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(4, weight=1)
        
        # Calendar file section
        self.calendar_frame = ctk.CTkFrame(self.main_frame)
        self.calendar_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.calendar_frame.grid_columnconfigure(1, weight=1)
        
        calendar_label = ctk.CTkLabel(self.calendar_frame, text="📅 Calendar File", 
                                     font=ctk.CTkFont(size=16, weight="bold"))
        calendar_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.calendar_entry = ctk.CTkEntry(self.calendar_frame, textvariable=self.calendar_path, 
                                          placeholder_text="Select calendar file...")
        self.calendar_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        calendar_btn = ctk.CTkButton(self.calendar_frame, text="Browse", 
                                    command=self.pick_calendar, width=100)
        calendar_btn.grid(row=0, column=2, padx=10, pady=10)
        
        # Office file section
        office_frame = ctk.CTkFrame(self.main_frame)
        office_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        office_frame.grid_columnconfigure(1, weight=1)
        
        office_label = ctk.CTkLabel(office_frame, text="📄 Office Format File", 
                                   font=ctk.CTkFont(size=16, weight="bold"))
        office_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.office_entry = ctk.CTkEntry(office_frame, textvariable=self.office_path, 
                                         placeholder_text="Select office format file...")
        self.office_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        office_btn = ctk.CTkButton(office_frame, text="Browse", 
                                  command=self.pick_office, width=100)
        office_btn.grid(row=0, column=2, padx=10, pady=10)
        
        # Output folder section
        output_frame = ctk.CTkFrame(self.main_frame)
        output_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        output_frame.grid_columnconfigure(1, weight=1)
        
        output_label = ctk.CTkLabel(output_frame, text="💾 Output Folder", 
                                   font=ctk.CTkFont(size=16, weight="bold"))
        output_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.output_entry = ctk.CTkEntry(output_frame, textvariable=self.output_path, 
                                        placeholder_text="Select output folder...")
        self.output_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        output_btn = ctk.CTkButton(output_frame, text="Browse", 
                                  command=self.pick_output_folder, width=100)
        output_btn.grid(row=0, column=2, padx=10, pady=10)
        
        # Start button
        start_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        start_frame.grid(row=3, column=0, pady=20)
        
        self.start_btn = ctk.CTkButton(start_frame, text="▶️ START PROCESSING", 
                                      command=self.run_processing,
                                      font=ctk.CTkFont(size=18, weight="bold"),
                                      height=50, width=300)
        self.start_btn.pack()
        
        # Progress bar
        self.progress = ctk.CTkProgressBar(self.main_frame)
        self.progress.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.progress.set(0)
        
        # Log section
        log_label = ctk.CTkLabel(self.main_frame, text="📋 Processing Log", 
                                font=ctk.CTkFont(size=16, weight="bold"))
        log_label.grid(row=5, column=0, padx=10, pady=(10, 5), sticky="w")
        
        self.log_textbox = ctk.CTkTextbox(self.main_frame, height=200)
        self.log_textbox.grid(row=6, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Initialize UI
        self.update_mode_ui()
        
    def create_footer(self):
        # Footer frame
        footer = ctk.CTkFrame(self.root, height=40)
        footer.grid(row=3, column=0, columnspan=2, sticky="ew")
        
        footer_label = ctk.CTkLabel(footer, text="© 2026 Sandeep | Cube Data Processor v2.0 | Settings stored in Windows Registry", 
                                   font=ctk.CTkFont(size=12))
        footer_label.pack(pady=10)
        
    def update_mode_ui(self):
        mode = self.mode_var.get()
        if mode == "date_only":
            self.calendar_frame.grid()
        else:
            self.calendar_frame.grid_remove()
            
    def update_grade_listbox(self):
        self.grade_listbox.delete("0.0", "end")
        for file in self.grade_files:
            self.grade_listbox.insert("end", f"📄 {os.path.basename(file)}\n")
            
    def add_grades(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
        for f in files:
            if f not in self.grade_files:
                self.grade_files.append(f)
        self.update_grade_listbox()
        
    def clear_grades(self):
        self.grade_files.clear()
        self.update_grade_listbox()
        
    def pick_office(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.office_path.set(path)
            
    def pick_calendar(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.calendar_path.set(path)
            
    def pick_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_path.set(folder)
            
    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.root.update_idletasks()
        
    def run_processing(self):
        # Validate inputs
        mode = self.mode_var.get()
        
        if mode in ["grade_only", "both"]:
            if not self.grade_files:
                messagebox.showerror("Error", "Please select grade files for grade processing.")
                return
        
        if mode in ["date_only", "both"]:
            if not self.calendar_path.get():
                messagebox.showerror("Error", "Please select calendar file for date processing.")
                return
        
        if not self.office_path.get():
            messagebox.showerror("Error", "Select office format file.")
            return

        if not self.output_path.get():
            messagebox.showerror("Error", "Select output folder.")
            return

        # Save settings to registry (NOT including office file path as requested)
        registry_settings.save_all_settings(self.grade_files, self.output_path.get(), self.calendar_path.get())

        # Clear log and start processing
        self.log_textbox.delete("0.0", "end")
        self.progress.set(0.3)
        self.root.update_idletasks()
        
        total = process_combined(
            self.grade_files,
            self.office_path.get(),
            self.output_path.get(),
            self.calendar_path.get(),
            mode,
            log_callback=self.log
        )
        
        self.progress.set(1.0)
        winsound.MessageBeep()
        messagebox.showinfo("✓ Completed", f"Processing Complete!\n\nTotal Operations: {total}")
        self.progress.set(0)
        
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = CubeDataProcessor()
    app.run()