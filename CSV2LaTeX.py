"""
Created on Wed Apr 23 11:07:17 2025

@author: hamedtabrizchi
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pyperclip  # For copying to clipboard

def select_file():
    file_path = filedialog.askopenfilename(
        title="Select a CSV File",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    if file_path:
        try:
            # Load the CSV file into a DataFrame
            df = pd.read_csv(file_path)
            # Display the DataFrame in the preview table
            display_table(df)
            # Store the DataFrame for LaTeX conversion
            app_data['dataframe'] = df
            app_data['file_path'] = file_path
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

def display_table(df):
    # Clear the treeview
    for item in tree.get_children():
        tree.delete(item)
    # Add columns
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    # Add rows
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

def generate_latex():
    if 'dataframe' not in app_data or app_data['dataframe'] is None:
        messagebox.showwarning("Warning", "No CSV file loaded!")
        return
    try:
        df = app_data['dataframe']
        caption = caption_entry.get()
        
        # Generate LaTeX code with caption
        latex_code = df.to_latex(index=False, escape=False, caption=caption)
        
        # Display the LaTeX code in the text box
        text_output.delete("1.0", tk.END)
        text_output.insert(tk.END, latex_code)
        # Store the LaTeX code for copying
        app_data['latex_code'] = latex_code
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate LaTeX code: {e}")

def copy_to_clipboard():
    if 'latex_code' in app_data and app_data['latex_code']:
        pyperclip.copy(app_data['latex_code'])
        messagebox.showinfo("Copied", "LaTeX code copied to clipboard!")
    else:
        messagebox.showwarning("Warning", "No LaTeX code to copy!")

# Initialize the main application window
app = tk.Tk()
app.title("CSV2LaTeX - CSV to LaTeX Table Converter")
app.geometry("900x700")

# Store application data
app_data = {}

# File selection frame
frame_file = ttk.Frame(app)
frame_file.pack(fill="x", padx=10, pady=10)
btn_select = ttk.Button(frame_file, text="Select CSV File", command=select_file)
btn_select.pack(side="left", padx=5)
lbl_file = ttk.Label(frame_file, text="No file selected")
lbl_file.pack(side="left", padx=5)

# Caption entry frame
frame_caption = ttk.Frame(app)
frame_caption.pack(fill="x", padx=10, pady=5)
lbl_caption = ttk.Label(frame_caption, text="Table Caption:")
lbl_caption.pack(side="left", padx=5)
caption_entry = ttk.Entry(frame_caption, width=50)
caption_entry.pack(side="left", padx=5)

# Table preview frame
frame_preview = ttk.LabelFrame(app, text="CSV Table Preview")
frame_preview.pack(fill="both", expand=True, padx=10, pady=10)
tree = ttk.Treeview(frame_preview)
tree.pack(side="left", fill="both", expand=True)

# Scrollbars for the table
scroll_y = ttk.Scrollbar(frame_preview, orient="vertical", command=tree.yview)
scroll_y.pack(side="right", fill="y")
scroll_x = ttk.Scrollbar(app, orient="horizontal", command=tree.xview)
scroll_x.pack(fill="x")
tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

# LaTeX output frame
frame_output = ttk.LabelFrame(app, text="LaTeX Table Code")
frame_output.pack(fill="both", expand=True, padx=10, pady=10)
text_output = tk.Text(frame_output, wrap="word", height=10)
text_output.pack(fill="both", expand=True, padx=5, pady=5)

# Buttons for generating LaTeX and copying to clipboard
frame_buttons = ttk.Frame(app)
frame_buttons.pack(fill="x", padx=10, pady=10)
btn_generate = ttk.Button(frame_buttons, text="Generate LaTeX Code", command=generate_latex)
btn_generate.pack(side="left", padx=5)
btn_copy = ttk.Button(frame_buttons, text="Copy to Clipboard", command=copy_to_clipboard)
btn_copy.pack(side="left", padx=5)

# Run the application
app.mainloop()