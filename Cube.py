import openpyxl
from openpyxl.drawing.image import Image as XLImage
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import winsound
from copy import copy

# Detect grade from filename
def extract_grade(filename):
    name = os.path.basename(filename).split('.')[0].upper()
    name = name.replace("_", ":").replace("-", ":")
    return name.strip()


# FILLED ROW CHECKER
def get_last_row(ws):
    row = 2
    while True:
        if ws.cell(row=row, column=2).value in (None, ""):
            return row - 1
        row += 1


# Copy images from source to target worksheet
def copy_images(source_ws, target_ws):
    if hasattr(source_ws, '_images') and source_ws._images:
        for img in source_ws._images:
            new_img = XLImage(img.ref)
            new_img.anchor = img.anchor
            target_ws.add_image(new_img)


# MAIN PROCESSING - SEPARATE MODE
def process_grade_separate(grade_file, office_file, output_folder, log):
    try:
        grade_wb = openpyxl.load_workbook(grade_file)
        grade_ws = grade_wb.active

        grade_name = extract_grade(grade_file)
        log(f"\n=== Processing {grade_file}")
        log(f"Detected Grade: {grade_name}")

        # Load office file
        office_wb = openpyxl.load_workbook(office_file)

        last_row = get_last_row(grade_ws)
        log(f"Total data rows: {last_row - 1}")

        # Get all sheets that match this grade
        matching_sheets = []
        for sheet_name in office_wb.sheetnames:
            ws = office_wb[sheet_name]
            b12 = str(ws["B12"].value).replace(" ", "").upper()
            if b12 == grade_name:
                matching_sheets.append(sheet_name)
        
        log(f"Found {len(matching_sheets)} sheets matching grade '{grade_name}'")
        
        if len(matching_sheets) == 0:
            log(f"⚠ WARNING: No sheets found with '{grade_name}' in cell B12!")
            return 0

        copy_count = 0
        sheet_index = 0

        # Loop through each data row
        for r in range(2, last_row + 1):
            if sheet_index >= len(matching_sheets):
                log(f"⚠ Warning: More data rows than matching sheets. Stopping at row {r}")
                break

            current_sheet_name = matching_sheets[sheet_index]
            ws = office_wb[current_sheet_name]
            
            # Read weight and strength values
            weight_values = [grade_ws.cell(row=r, column=c).value for c in range(2, 8)]
            strength_values = [grade_ws.cell(row=r, column=c).value for c in range(9, 15)]

            # Write values
            for i, v in enumerate(weight_values):
                ws.cell(row=25, column=3 + i, value=v)
            for i, v in enumerate(strength_values):
                ws.cell(row=27, column=3 + i, value=v)

            copy_count += 1
            log(f"✓ Row {r} → Sheet: {current_sheet_name}")
            sheet_index += 1

        # Save separate file
        base = os.path.basename(office_file).split(".")[0]
        outname = f"{base}_{grade_name}_Processed.xlsx"
        outpath = os.path.join(output_folder, outname)

        office_wb.save(outpath)
        log(f"✓ Saved → {outpath}")

        return copy_count

    except Exception as e:
        log(f"✖ ERROR: {e}")
        return 0


# MAIN PROCESSING - COMBINE MODE
def process_all_grades_combined(grade_files, office_file, output_folder, log):
    try:
        log(f"\n=== COMBINE MODE: Processing all grades into one file ===")
        
        # Load office file ONCE
        office_wb = openpyxl.load_workbook(office_file)
        
        all_matching_sheets = []
        total_copy_count = 0
        
        # Process each grade file
        for grade_file in grade_files:
            grade_wb = openpyxl.load_workbook(grade_file)
            grade_ws = grade_wb.active
            grade_name = extract_grade(grade_file)
            
            log(f"\n--- Processing Grade: {grade_name} ---")
            
            last_row = get_last_row(grade_ws)
            log(f"Data rows: {last_row - 1}")
            
            # Find matching sheets for this grade
            matching_sheets = []
            for sheet_name in office_wb.sheetnames:
                ws = office_wb[sheet_name]
                b12 = str(ws["B12"].value).replace(" ", "").upper()
                if b12 == grade_name:
                    matching_sheets.append(sheet_name)
            
            log(f"Found {len(matching_sheets)} sheets for '{grade_name}'")
            
            if len(matching_sheets) == 0:
                log(f"⚠ No sheets found for '{grade_name}'")
                continue
            
            sheet_index = 0
            
            # Copy data for this grade
            for r in range(2, last_row + 1):
                if sheet_index >= len(matching_sheets):
                    log(f"⚠ More rows than sheets for {grade_name}")
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
                log(f"✓ {grade_name} Row {r} → {current_sheet_name}")
                sheet_index += 1
        
        # Save ONE combined file
        base = os.path.basename(office_file).split(".")[0]
        outname = f"{base}_ALL_GRADES_Combined.xlsx"
        outpath = os.path.join(output_folder, outname)
        
        office_wb.save(outpath)
        log(f"\n✓✓✓ COMBINED FILE SAVED → {outpath}")
        
        return total_copy_count
        
    except Exception as e:
        log(f"✖ ERROR: {e}")
        return 0


# ------------- GUI LOGIC -------------

def run_processing():
    if not grade_files:
        messagebox.showerror("Error", "Please select grade files.")
        return

    if not office_path.get():
        messagebox.showerror("Error", "Select office format file.")
        return

    if not output_path.get():
        messagebox.showerror("Error", "Select output folder.")
        return

    log_box.delete("1.0", "end")
    total = 0

    if mode_var.get() == 2:  # COMBINE MODE
        progress["value"] = 50
        root.update_idletasks()
        
        total = process_all_grades_combined(
            grade_files,
            office_path.get(),
            output_path.get(),
            log=lambda m: log_box.insert(tk.END, m + "\n")
        )
        
        progress["value"] = 100
        
    else:  # SEPARATE MODE
        for i, file in enumerate(grade_files):
            progress["value"] = (i + 1) / len(grade_files) * 100
            root.update_idletasks()

            total += process_grade_separate(
                file,
                office_path.get(),
                output_path.get(),
                log=lambda m: log_box.insert(tk.END, m + "\n")
            )

    winsound.MessageBeep()
    messagebox.showinfo("✓ Completed", f"Processing Complete!\n\nTotal Rows Copied: {total}")


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


def pick_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_path.set(folder)


# ------------------- MODERN GUI -------------------

root = tk.Tk()
root.title("Cube Data Processor")
root.geometry("850x750")
root.configure(bg="#f0f0f0")

grade_files = []
office_path = tk.StringVar()
output_path = tk.StringVar()
mode_var = tk.IntVar(value=1)

# Style configuration
style = ttk.Style()
style.theme_use('clam')
style.configure('TButton', padding=6, relief="flat", background="#0078d4", foreground="white")
style.map('TButton', background=[('active', '#005a9e')])

# Header
header_frame = tk.Frame(root, bg="#0078d4", height=70)
header_frame.pack(fill=tk.X)
header_frame.pack_propagate(False)

title_label = tk.Label(header_frame, text="🔷 Cube Data Processor", 
                       font=("Segoe UI", 18, "bold"), bg="#0078d4", fg="white")
title_label.pack(pady=20)

# Main container
main_frame = tk.Frame(root, bg="#f0f0f0", padx=20, pady=20)
main_frame.pack(fill=tk.BOTH, expand=True)

# Grade Files Section
grade_frame = tk.LabelFrame(main_frame, text="📁 Grade Files", font=("Segoe UI", 10, "bold"),
                            bg="#f0f0f0", padx=10, pady=10)
grade_frame.pack(fill=tk.BOTH, expand=True, pady=5)

btn_frame = tk.Frame(grade_frame, bg="#f0f0f0")
btn_frame.pack(fill=tk.X, pady=5)

ttk.Button(btn_frame, text="➕ Add Files", command=add_grades).pack(side=tk.LEFT, padx=5)
ttk.Button(btn_frame, text="🗑️ Clear All", command=clear_grades).pack(side=tk.LEFT, padx=5)

grade_listbox = tk.Listbox(grade_frame, height=5, font=("Consolas", 9), 
                           bg="white", relief=tk.FLAT, bd=1)
grade_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

# Office File Section
office_frame = tk.LabelFrame(main_frame, text="📄 Office Format File", 
                            font=("Segoe UI", 10, "bold"), bg="#f0f0f0", padx=10, pady=10)
office_frame.pack(fill=tk.X, pady=5)

ttk.Button(office_frame, text="Select File", command=pick_office).pack(anchor=tk.W, pady=5)
office_entry = tk.Entry(office_frame, textvariable=office_path, font=("Segoe UI", 9),
                       bg="white", relief=tk.FLAT, bd=1)
office_entry.pack(fill=tk.X, ipady=5)

# Output Folder Section
output_frame = tk.LabelFrame(main_frame, text="💾 Output Folder", 
                            font=("Segoe UI", 10, "bold"), bg="#f0f0f0", padx=10, pady=10)
output_frame.pack(fill=tk.X, pady=5)

ttk.Button(output_frame, text="Select Folder", command=pick_output_folder).pack(anchor=tk.W, pady=5)
output_entry = tk.Entry(output_frame, textvariable=output_path, font=("Segoe UI", 9),
                       bg="white", relief=tk.FLAT, bd=1)
output_entry.pack(fill=tk.X, ipady=5)

# Processing Mode
mode_frame = tk.LabelFrame(main_frame, text="⚙️ Processing Mode", 
                          font=("Segoe UI", 10, "bold"), bg="#f0f0f0", padx=10, pady=10)
mode_frame.pack(fill=tk.X, pady=5)

tk.Radiobutton(mode_frame, text="📑 Separate Files (One per grade)", 
               variable=mode_var, value=1, font=("Segoe UI", 9),
               bg="#f0f0f0", activebackground="#f0f0f0").pack(anchor=tk.W)
tk.Radiobutton(mode_frame, text="📦 Combined File (All grades in one)", 
               variable=mode_var, value=2, font=("Segoe UI", 9),
               bg="#f0f0f0", activebackground="#f0f0f0").pack(anchor=tk.W)

# Start Button
start_btn = tk.Button(main_frame, text="▶️  START PROCESSING", command=run_processing,
                     font=("Segoe UI", 11, "bold"), bg="#28a745", fg="white",
                     activebackground="#218838", relief=tk.FLAT, cursor="hand2",
                     padx=20, pady=10)
start_btn.pack(pady=15)

# Progress Bar
progress = ttk.Progressbar(main_frame, length=400, mode="determinate")
progress.pack(pady=5)

# Log Section
log_frame = tk.LabelFrame(main_frame, text="📋 Processing Log", 
                         font=("Segoe UI", 10, "bold"), bg="#f0f0f0", padx=10, pady=10)
log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

log_box = tk.Text(log_frame, height=12, font=("Consolas", 8), bg="white",
                 relief=tk.FLAT, bd=1, wrap=tk.WORD)
log_box.pack(fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(log_box, command=log_box.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_box.config(yscrollcommand=scrollbar.set)

root.mainloop()