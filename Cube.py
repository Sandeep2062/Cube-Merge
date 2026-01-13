# Start Button - More space above and below
start_btn_container = tk.Frame(mainimport openpyxl
from openpyxl.drawing.image import Image as XLImage
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import winsound
from copy import deepcopy
import shutil
import json

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

# Settings file location (same folder as EXE)
def get_settings_path():
    """Get settings file path in same folder as script/EXE"""
    if getattr(sys, 'frozen', False):
        # Running as EXE
        base_path = os.path.dirname(sys.executable)
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, "cube_settings.json")

SETTINGS_FILE = get_settings_path()

# Load saved settings
def load_settings():
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                # Convert grade_files list to filenames only
                if "grade_files" in settings:
                    grade_files_data = settings["grade_files"]
                else:
                    grade_files_data = []
                return settings, grade_files_data
    except:
        pass
    return {"output_path": "", "calendar_path": ""}, []

# Save settings (remembers grade files and paths except office)
def save_settings(grade_file_list, output, calendar):
    try:
        settings = {
            "grade_files": grade_file_list,  # Save full paths
            "output_path": output,
            "calendar_path": calendar
        }
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        print(f"Could not save settings: {e}")

# Smart grade extraction - handles M20, M15, and Mortar_1_4 format
def extract_grade(filename):
    name = os.path.basename(filename).split('.')[0].upper()
    
    # Check if it's mortar format (contains underscore followed by numbers)
    # E.g., MORTAR_1_4 → 1:4
    if "MORTAR" in name and "_" in name:
        parts = name.split("_")
        # Get the ratio part after MORTAR_
        if len(parts) >= 3:  # MORTAR_1_4
            ratio = f"{parts[-2]}:{parts[-1]}"
            return ratio
    
    # For regular grades (M20, M15, etc.) - just clean up
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
def load_calendar_data(calendar_file, log):
    """Load calendar dates from Excel file"""
    try:
        if not calendar_file or not os.path.exists(calendar_file):
            log("⚠ No calendar file selected")
            return None
        
        wb = load_workbook_safe(calendar_file)
        ws = wb.active
        
        calendar_dict = {}
        row = 2  # Start from row 2
        
        while True:
            casting_date = ws.cell(row=row, column=1).value  # Column A
            if not casting_date:
                break
            
            date_7 = ws.cell(row=row, column=2).value   # Column B (7 days)
            date_28 = ws.cell(row=row, column=3).value  # Column C (28 days)
            
            # Store as string for matching
            if casting_date:
                date_str = str(casting_date).strip()
                calendar_dict[date_str] = {
                    "7_days": str(date_7).strip() if date_7 else "",
                    "28_days": str(date_28).strip() if date_28 else ""
                }
            
            row += 1
        
        wb.close()
        log(f"✓ Calendar loaded: {len(calendar_dict)} dates")
        return calendar_dict
        
    except Exception as e:
        log(f"✖ Calendar load error: {e}")
        return None

# PROCESS WITH GRADE AND/OR DATE
def process_combined(grade_files, office_file, output_folder, calendar_file, mode, log):
    try:
        log(f"\n{'='*60}")
        log(f"PROCESSING MODE: {mode.upper().replace('_', ' ')}")
        log(f"{'='*60}")
        
        # Load calendar if date mode
        calendar_data = None
        if mode in ["date_only", "both"]:
            calendar_data = load_calendar_data(calendar_file, log)
            if not calendar_data:
                log("✖ Cannot proceed without calendar file")
                return 0
        
        # Create output filename (NO TIMESTAMP)
        base = os.path.basename(office_file).split(".")[0]
        outname = f"{base}_Processed.xlsx"
        outpath = os.path.join(output_folder, outname)
        
        # Copy template to preserve images
        shutil.copy2(office_file, outpath)
        
        # Load and modify the copy
        office_wb = load_workbook_safe(outpath)
        
        total_copy_count = 0
        
        # GRADE PROCESSING
        if mode in ["grade_only", "both"] and grade_files:
            log(f"\n--- GRADE PROCESSING ---")
            
            for grade_file in grade_files:
                grade_wb = load_workbook_safe(grade_file)
                grade_ws = grade_wb.active
                grade_name = extract_grade(grade_file)
                
                log(f"\nProcessing: {os.path.basename(grade_file)}")
                log(f"Looking for grade: {grade_name}")
                
                last_row = get_last_row(grade_ws)
                log(f"Data rows: {last_row - 1}")
                
                # Find matching sheets - normalize both sides for comparison
                matching_sheets = []
                for sheet_name in office_wb.sheetnames:
                    ws = office_wb[sheet_name]
                    b12_value = ws["B12"].value
                    if b12_value:
                        # Normalize B12 value
                        b12 = str(b12_value).replace(" ", "").upper()
                        
                        # For comparison, also normalize grade_name
                        grade_normalized = grade_name.replace(" ", "").upper()
                        
                        if b12 == grade_normalized:
                            matching_sheets.append(sheet_name)
                            log(f"  ✓ Matched sheet: {sheet_name} (B12={b12})")
                
                log(f"Total matching sheets: {len(matching_sheets)}")
                
                if len(matching_sheets) == 0:
                    log(f"⚠ No sheets found with B12='{grade_name}'")
                    grade_wb.close()
                    continue
                
                sheet_index = 0
                
                # Copy grade data
                for r in range(2, last_row + 1):
                    if sheet_index >= len(matching_sheets):
                        log(f"⚠ More data rows than available sheets")
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
                    log(f"  ✓ Row {r} → {current_sheet_name}")
                    sheet_index += 1
                
                grade_wb.close()
        
        # DATE PROCESSING
        if mode in ["date_only", "both"] and calendar_data:
            log(f"\n--- DATE PROCESSING ---")
            
            updated_count = 0
            
            for sheet_name in office_wb.sheetnames:
                ws = office_wb[sheet_name]
                
                # Read casting date from C17
                casting_date_cell = ws["C17"].value
                if not casting_date_cell:
                    continue
                
                casting_date = str(casting_date_cell).strip()
                
                # Look up in calendar
                if casting_date in calendar_data:
                    date_7 = calendar_data[casting_date]["7_days"]
                    date_28 = calendar_data[casting_date]["28_days"]
                    
                    # Write 7-day date to C18
                    if date_7:
                        ws["C18"] = date_7
                    
                    # Write 28-day date to F18
                    if date_28:
                        ws["F18"] = date_28
                    
                    updated_count += 1
                    log(f"✓ {sheet_name}: {casting_date} → 7d:{date_7}, 28d:{date_28}")
                else:
                    log(f"⚠ Date not in calendar: {casting_date} ({sheet_name})")
            
            log(f"\nSheets updated: {updated_count}")
        
        # Save combined file
        office_wb.save(outpath)
        office_wb.close()
        
        log(f"\n{'='*60}")
        log(f"✓✓✓ SAVED: {outpath}")
        log(f"{'='*60}")
        
        return total_copy_count
        
    except Exception as e:
        log(f"✖ ERROR: {e}")
        import traceback
        log(traceback.format_exc())
        return 0

# ------------- GUI LOGIC -------------

def run_processing():
    # Validate inputs based on mode
    mode = mode_var.get()
    
    if mode in ["grade_only", "both"]:
        if not grade_files:
            messagebox.showerror("Error", "Please select grade files for grade processing.")
            return
    
    if mode in ["date_only", "both"]:
        if not calendar_path.get():
            messagebox.showerror("Error", "Please select calendar file for date processing.")
            return
    
    if not office_path.get():
        messagebox.showerror("Error", "Select office format file.")
        return

    if not output_path.get():
        messagebox.showerror("Error", "Select output folder.")
        return

    # Save settings (grade files, output path, calendar - NOT office)
    save_settings(grade_files, output_path.get(), calendar_path.get())

    log_box.delete("1.0", "end")
    
    progress["value"] = 50
    root.update_idletasks()
    
    total = process_combined(
        grade_files,
        office_path.get(),
        output_path.get(),
        calendar_path.get(),
        mode,
        log=lambda m: log_box.insert(tk.END, m + "\n")
    )
    
    progress["value"] = 100
    
    winsound.MessageBeep()
    messagebox.showinfo("✓ Completed", f"Processing Complete!\n\nTotal Operations: {total}")

def add_grades():
    files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
    for f in files:
        if f not in grade_files:
            grade_files.append(f)
            grade_listbox.insert(tk.END, os.path.basename(f))

def clear_grades():
    grade_files.clear()
    grade_listbox.delete(0, tk.END)

def pick_office():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path:
        office_path.set(path)

def pick_calendar():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path:
        calendar_path.set(path)

def pick_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_path.set(folder)

def update_mode_ui():
    """Show/hide sections based on mode"""
    mode = mode_var.get()
    
    if mode == "grade_only":
        grade_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10), before=office_frame)
        calendar_frame.pack_forget()
    elif mode == "date_only":
        grade_frame.pack_forget()
        calendar_frame.pack(fill=tk.X, pady=(0, 10), before=office_frame)
    else:  # both
        grade_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10), before=office_frame)
        calendar_frame.pack(fill=tk.X, pady=(0, 10), before=office_frame)

# ------------------- ENHANCED DARK UI V6.1 -------------------

root = tk.Tk()
root.title("Cube Data Processor")
root.geometry("1000x950")  # Bigger window
root.minsize(950, 900)  # Minimum size to prevent crushing
root.configure(bg="#0f0f0f")

try:
    root.iconbitmap("icon.ico")
except:
    pass

# Load saved settings
import sys
settings, saved_grade_files = load_settings()

grade_files = []
office_path = tk.StringVar()
output_path = tk.StringVar(value=settings.get("output_path", ""))
calendar_path = tk.StringVar(value=settings.get("calendar_path", ""))
mode_var = tk.StringVar(value="both")

# Enhanced Dark Colors
BG_DARK = "#0f0f0f"
BG_CARD = "#1a1a1a"
BG_CARD_HOVER = "#252525"
BG_INPUT = "#242424"
TEXT_PRIMARY = "#ffffff"
TEXT_SECONDARY = "#a0a0a0"
ACCENT_TEAL = "#0d9488"
ACCENT_GREEN = "#10b981"
BORDER_COLOR = "#2a2a2a"
HOVER_COLOR = "#1e3a3a"

# Styles
style = ttk.Style()
style.theme_use('clam')

style.configure('Dark.TButton', padding=9, relief="flat", background=ACCENT_TEAL, 
                foreground="white", font=("Segoe UI", 9), borderwidth=0)
style.map('Dark.TButton', background=[('active', '#0a6b62')])

style.configure('Action.TButton', padding=14, background=ACCENT_GREEN, 
                foreground="white", font=("Segoe UI", 12, "bold"), borderwidth=0)
style.map('Action.TButton', background=[('active', '#059669')])

style.configure("Dark.Horizontal.TProgressbar", background=ACCENT_GREEN, 
                troughcolor=BG_INPUT, borderwidth=0, thickness=8)

# GRADIENT HEADER
header_frame = tk.Frame(root, bg="#0d9488", height=95)
header_frame.pack(fill=tk.X)
header_frame.pack_propagate(False)

# Logo
logo_container = tk.Frame(header_frame, bg="#0d9488")
logo_container.place(x=30, y=17)

try:
    from PIL import Image, ImageTk
    logo_img = Image.open("logo.png")
    logo_img = logo_img.resize((60, 60), Image.Resampling.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(logo_container, image=logo_photo, bg="#0d9488")
    logo_label.image = logo_photo
    logo_label.pack()
except:
    logo_label = tk.Label(logo_container, text="🔷", font=("Segoe UI", 44), bg="#0d9488", fg="white")
    logo_label.pack()

# Title
title_container = tk.Frame(header_frame, bg="#0d9488")
title_container.pack(expand=True)

title_label = tk.Label(title_container, text="CUBE DATA PROCESSOR", 
                       font=("Segoe UI", 26, "bold"), bg="#0d9488", fg="white", 
                       pady=12)
title_label.pack()

subtitle_label = tk.Label(title_container, text="Professional Edition", 
                         font=("Segoe UI", 9), bg="#0d9488", fg="#d1fae5")
subtitle_label.pack()

# Developer Credit
credit_frame = tk.Frame(header_frame, bg="#0d9488")
credit_frame.place(relx=1.0, y=20, anchor="ne", x=-30)

credit_label = tk.Label(credit_frame, text="Developed by", 
                       font=("Segoe UI", 8), bg="#0d9488", fg="#d1fae5")
credit_label.pack()

dev_name_label = tk.Label(credit_frame, text="SANDEEP", 
                         font=("Segoe UI", 12, "bold"), bg="#0d9488", fg="white")
dev_name_label.pack()

github_label = tk.Label(credit_frame, text="github.com/Sandeep2062", 
                       font=("Segoe UI", 8), bg="#0d9488", fg="#5eead4", cursor="hand2", 
                       underline=True)
github_label.pack()
github_label.bind("<Button-1>", lambda e: os.system("start https://github.com/Sandeep2062/Cube-Data-Processor"))

# Main Container with Scrollbar
main_canvas = tk.Canvas(root, bg=BG_DARK, highlightthickness=0)
main_scrollbar = tk.Scrollbar(root, orient="vertical", command=main_canvas.yview)
main_container = tk.Frame(main_canvas, bg=BG_DARK)

main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=25, pady=20)
main_canvas.create_window((0, 0), window=main_container, anchor="nw")

def on_frame_configure(event):
    main_canvas.configure(scrollregion=main_canvas.bbox("all"))

main_container.bind("<Configure>", on_frame_configure)

# Mouse wheel scrolling
def on_mousewheel(event):
    main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

main_canvas.bind_all("<MouseWheel>", on_mousewheel)

# Processing Mode Selection
mode_selection_frame = tk.LabelFrame(main_container, text="  ⚙️ PROCESSING MODE  ", 
                                    font=("Segoe UI", 11, "bold"), bg=BG_CARD, 
                                    fg=TEXT_PRIMARY, bd=0, relief=tk.FLAT, 
                                    padx=20, pady=18, highlightbackground=BORDER_COLOR, 
                                    highlightthickness=1)
mode_selection_frame.pack(fill=tk.X, pady=(0, 10))

tk.Radiobutton(mode_selection_frame, text="📊 Grade Only", 
               variable=mode_var, value="grade_only", font=("Segoe UI", 10),
               bg=BG_CARD, fg=TEXT_PRIMARY, activebackground=BG_CARD, 
               activeforeground=TEXT_PRIMARY, selectcolor=BG_INPUT,
               command=update_mode_ui, cursor="hand2").pack(anchor=tk.W, pady=4)

tk.Radiobutton(mode_selection_frame, text="📅 Date Only", 
               variable=mode_var, value="date_only", font=("Segoe UI", 10),
               bg=BG_CARD, fg=TEXT_PRIMARY, activebackground=BG_CARD, 
               activeforeground=TEXT_PRIMARY, selectcolor=BG_INPUT,
               command=update_mode_ui, cursor="hand2").pack(anchor=tk.W, pady=4)

tk.Radiobutton(mode_selection_frame, text="🔄 Both (Grade + Date)", 
               variable=mode_var, value="both", font=("Segoe UI", 10),
               bg=BG_CARD, fg=TEXT_PRIMARY, activebackground=BG_CARD, 
               activeforeground=TEXT_PRIMARY, selectcolor=BG_INPUT,
               command=update_mode_ui, cursor="hand2").pack(anchor=tk.W, pady=4)

# Grade Files Section
grade_frame = tk.LabelFrame(main_container, text="  📁 GRADE FILES  ", 
                            font=("Segoe UI", 11, "bold"), bg=BG_CARD, 
                            fg=TEXT_PRIMARY, bd=0, relief=tk.FLAT, 
                            padx=20, pady=18, highlightbackground=BORDER_COLOR,
                            highlightthickness=1)

btn_frame = tk.Frame(grade_frame, bg=BG_CARD)
btn_frame.pack(fill=tk.X, pady=(0, 12))

ttk.Button(btn_frame, text="➕ Add Files", command=add_grades, style='Dark.TButton').pack(side=tk.LEFT, padx=5)
ttk.Button(btn_frame, text="🗑️ Clear", command=clear_grades, style='Dark.TButton').pack(side=tk.LEFT, padx=5)

grade_listbox = tk.Listbox(grade_frame, height=3, font=("Consolas", 9), 
                           bg=BG_INPUT, fg=TEXT_PRIMARY, relief=tk.FLAT, bd=0, 
                           highlightthickness=1, highlightbackground=BORDER_COLOR,
                           selectbackground=ACCENT_TEAL, selectforeground="white")
grade_listbox.pack(fill=tk.BOTH, expand=True)

# Load saved grade files
for gf in saved_grade_files:
    if os.path.exists(gf):
        grade_files.append(gf)
        grade_listbox.insert(tk.END, os.path.basename(gf))

# Calendar File Section
calendar_frame = tk.LabelFrame(main_container, text="  📅 CALENDAR FILE  ", 
                              font=("Segoe UI", 11, "bold"), bg=BG_CARD, 
                              fg=TEXT_PRIMARY, bd=0, relief=tk.FLAT, 
                              padx=20, pady=18, highlightbackground=BORDER_COLOR,
                              highlightthickness=1)

ttk.Button(calendar_frame, text="📂 Select Calendar", command=pick_calendar, 
          style='Dark.TButton').pack(anchor=tk.W, pady=(0, 10))
calendar_entry = tk.Entry(calendar_frame, textvariable=calendar_path, font=("Segoe UI", 9),
                         bg=BG_INPUT, fg=TEXT_PRIMARY, relief=tk.FLAT, bd=0, 
                         insertbackground="white", highlightthickness=1,
                         highlightbackground=BORDER_COLOR)
calendar_entry.pack(fill=tk.X, ipady=12, padx=2)

# Office File Section
office_frame = tk.LabelFrame(main_container, text="  📄 OFFICE FORMAT FILE  ", 
                            font=("Segoe UI", 11, "bold"), bg=BG_CARD, 
                            fg=TEXT_PRIMARY, bd=0, relief=tk.FLAT, 
                            padx=20, pady=18, highlightbackground=BORDER_COLOR,
                            highlightthickness=1)
office_frame.pack(fill=tk.X, pady=(0, 10))

ttk.Button(office_frame, text="📂 Select File", command=pick_office, 
          style='Dark.TButton').pack(anchor=tk.W, pady=(0, 10))
office_entry = tk.Entry(office_frame, textvariable=office_path, font=("Segoe UI", 9),
                       bg=BG_INPUT, fg=TEXT_PRIMARY, relief=tk.FLAT, bd=0, 
                       insertbackground="white", highlightthickness=1,
                       highlightbackground=BORDER_COLOR)
office_entry.pack(fill=tk.X, ipady=12, padx=2)

# Output Folder Section
output_frame = tk.LabelFrame(main_container, text="  💾 OUTPUT FOLDER  ", 
                            font=("Segoe UI", 11, "bold"), bg=BG_CARD, 
                            fg=TEXT_PRIMARY, bd=0, relief=tk.FLAT, 
                            padx=20, pady=18, highlightbackground=BORDER_COLOR,
                            highlightthickness=1)
output_frame.pack(fill=tk.X, pady=(0, 10))

ttk.Button(output_frame, text="📂 Select Folder", command=pick_output_folder, 
          style='Dark.TButton').pack(anchor=tk.W, pady=(0, 10))
output_entry = tk.Entry(output_frame, textvariable=output_path, font=("Segoe UI", 9),
                       bg=BG_INPUT, fg=TEXT_PRIMARY, relief=tk.FLAT, bd=0, 
                       insertbackground="white", highlightthickness=1,
                       highlightbackground=BORDER_COLOR)
output_entry.pack(fill=tk.X, ipady=12, padx=2)

# Start Button - Fixed padding for proper display
start_btn = tk.Button(main_container, text="▶️  START PROCESSING", command=run_processing,
                     font=("Segoe UI", 13, "bold"), bg=ACCENT_GREEN, fg="white",
                     activebackground="#059669", relief=tk.FLAT, cursor="hand2",
                     padx=40, pady=16, borderwidth=0, height=2)
start_btn.pack(pady=(20, 18))

# Progress Bar
progress_frame = tk.Frame(main_container, bg=BG_DARK)
progress_frame.pack(fill=tk.X, pady=(0, 12))

progress = ttk.Progressbar(progress_frame, length=700, mode="determinate", 
                          style="Dark.Horizontal.TProgressbar")
progress.pack()

# Log Section
log_frame = tk.LabelFrame(main_container, text="  📋 PROCESSING LOG  ", 
                         font=("Segoe UI", 11, "bold"), bg=BG_CARD, 
                         fg=TEXT_PRIMARY, bd=0, relief=tk.FLAT, 
                         padx=20, pady=18, highlightbackground=BORDER_COLOR,
                         highlightthickness=1)
log_frame.pack(fill=tk.BOTH, expand=True)

log_scrollbar = tk.Scrollbar(log_frame, bg=BG_INPUT)
log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

log_box = tk.Text(log_frame, height=10, font=("Consolas", 9), bg=BG_INPUT, 
                 fg="#6ee7b7", relief=tk.FLAT, bd=0, wrap=tk.WORD, 
                 yscrollcommand=log_scrollbar.set, insertbackground="white")
log_box.pack(fill=tk.BOTH, expand=True)
log_scrollbar.config(command=log_box.yview)

# Footer
footer = tk.Label(root, text="© 2026 Sandeep | github.com/Sandeep2062/Cube-Data-Processor", 
                 font=("Segoe UI", 8), bg=BG_DARK, fg=TEXT_SECONDARY)
footer.pack(pady=10)

# Initialize UI
update_mode_ui()

root.mainloop()