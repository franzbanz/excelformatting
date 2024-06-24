import tkinter as tk
from tkinter import filedialog
import os
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

def parse_cell(cell):
    try:
        # Use regex to find all float numbers within the cell content
        points = re.findall(r"[-+]?\d*\.\d+|\d+", cell)
        return [float(point) for point in points]
    except Exception as e:
        print(f"Error parsing cell: {cell}, error: {e}")
        return []

def browse_file(entry, save=False):
    func = filedialog.asksaveasfilename if save else filedialog.askopenfilename
    filetypes = [("Excel files", "*.xlsx")]
    filename = func(title="Select File", defaultextension=".xlsx", filetypes=filetypes)
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def browse_input_file():
    browse_file(input_entry)

def browse_output_file():
    browse_file(output_entry, save=True)

def process_files():
    input_file = input_entry.get()
    output_file = output_entry.get()
    if not input_file or not os.path.isfile(input_file):
        result_label.config(text="Invalid input file path")
        return
    if not output_file:
        result_label.config(text="Invalid output file path")
        return
    
    try:
        # Load the workbook and select the active worksheet
        wb = load_workbook(input_file)
        ws = wb.active

        # Create a new workbook for the expanded data
        new_wb = Workbook()
        new_ws = new_wb.active

        # Copy the first two rows directly to the new worksheet
        for row in ws.iter_rows(min_row=1, max_row=2, values_only=True):
            new_ws.append(row)

        # Iterate over each cell starting from the third row
        for r, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
            new_rows = [[] for _ in range(4)]  # Assuming each cell has at most 4 data points
            for c, cell in enumerate(row):
                points = parse_cell(str(cell))
                for i, point in enumerate(points):
                    if i < 4:
                        new_rows[i].extend([None] * (c - len(new_rows[i])))
                        new_rows[i].append(point)
            
            max_len = max(len(row) for row in new_rows)
            for row in new_rows:
                row.extend([None] * (max_len - len(row)))
                new_ws.append(row)

        # Save the expanded data to a new Excel file
        new_wb.save(output_file)

        # Load the new workbook to apply formatting
        wb = load_workbook(output_file)
        ws = wb.active

        # Define the colors
        blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Apply alternating colors starting from the third row
        for r in range(2, ws.max_row + 1):
            fill = blue_fill if ((r - 2) // 4) % 2 == 0 else green_fill
            for cell in ws[r]:
                cell.fill = fill

        wb.save(output_file)
        result_label.config(text="Processing completed successfully")
    except Exception as e:
        result_label.config(text=f"Error: {e}")

# Create the main window
root = tk.Tk()
root.title("File Processor")

# Create and place the widgets
tk.Label(root, text="Input File").grid(row=0, column=0, padx=10, pady=10)
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=browse_input_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Output File").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=browse_output_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Process", command=process_files).grid(row=2, column=0, columnspan=3, pady=20)
result_label = tk.Label(root, text="")
result_label.grid(row=3, column=0, columnspan=3, pady=10)

# Run the main event loop
root.mainloop()
