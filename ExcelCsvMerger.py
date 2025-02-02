import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def select_input_folder():
    folder_selected = filedialog.askdirectory()
    input_folder_path.set(folder_selected)

def select_output_folder():
    folder_selected = filedialog.askdirectory()
    output_folder_path.set(folder_selected)

def merge_files():
    input_folder = input_folder_path.get()
    output_folder = output_folder_path.get()

    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both input and output folders")
        return

    all_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith('.xlsx') or f.endswith('.csv')]
    if not all_files:
        messagebox.showerror("Error", "No Excel or CSV files found in the selected folder")
        return

    combined_df = pd.DataFrame()

    for file in all_files:
        if file.endswith('.xlsx'):
            df = pd.read_excel(file)
        elif file.endswith('.csv'):
            df = pd.read_csv(file)
        combined_df = pd.concat([combined_df, df], ignore_index=True)

    output_file = os.path.join(output_folder, 'combined.csv')
    combined_df.to_csv(output_file, index=False)


    messagebox.showinfo("Success", f"Files have been merged and saved to {output_file}")

# Setup the GUI
root = tk.Tk()
root.title("Excel and CSV Merger")

input_folder_path = tk.StringVar()
output_folder_path = tk.StringVar()

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

tk.Label(frame, text="Select Input Folder:").grid(row=0, column=0, sticky="w")
tk.Entry(frame, textvariable=input_folder_path, width=50).grid(row=0, column=1)
tk.Button(frame, text="Browse", command=select_input_folder).grid(row=0, column=2)

tk.Label(frame, text="Select Output Folder:").grid(row=1, column=0, sticky="w")
tk.Entry(frame, textvariable=output_folder_path, width=50).grid(row=1, column=1)
tk.Button(frame, text="Browse", command=select_output_folder).grid(row=1, column=2)

tk.Button(frame, text="Merge Files", command=merge_files).grid(row=2, columnspan=3, pady=10)

root.mainloop()
