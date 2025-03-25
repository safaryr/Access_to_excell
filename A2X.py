
#!/usr/bin/env python3
import os
import subprocess
import pandas as pd
from tkinter import Tk, filedialog, Button, Label, StringVar, OptionMenu, messagebox

# Function to list all tables in the Access database
def list_tables(database_path):
    try:
        result = subprocess.run(['mdb-tables', '-1', database_path], stdout=subprocess.PIPE, text=True)
        tables = result.stdout.strip().split('\n')
        return [table for table in tables if table]
    except Exception as e:
        messagebox.showerror("Error", f"Error listing tables: {e}")
        return []

# Function to export a specific table to an Excel file
def export_table_to_excel(database_path, table_name, output_file):
    try:
        # Export table as CSV using mdb-export
        csv_file = f"{table_name}.csv"
        with open(csv_file, 'w') as f:
            subprocess.run(['mdb-export', database_path, table_name], stdout=f)
        
        # Read the CSV into a pandas DataFrame
        df = pd.read_csv(csv_file)
        
        # Export DataFrame to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=table_name, index=False)
        
        # Clean up temporary CSV file
        os.remove(csv_file)
        messagebox.showinfo("Success", f"Table '{table_name}' exported successfully to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error exporting table '{table_name}': {e}")

# Function to handle file selection
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Access Files", "*.mdb *.accdb")])
    if file_path:
        database_path.set(file_path)
        tables = list_tables(file_path)
        if tables:
            table_var.set(tables[0])  # Set the first table as default
            table_menu['menu'].delete(0, 'end')  # Clear current options in the dropdown menu
            for table in tables:
                table_menu['menu'].add_command(label=table, command=lambda value=table: table_var.set(value))
        else:
            messagebox.showerror("Error", "No tables found in the selected database.")

# Function to handle export operation
def export_data():
    if not database_path.get():
        messagebox.showerror("Error", "Please select an Access database first.")
        return
    
    if not table_var.get():
        messagebox.showerror("Error", "Please select a table.")
        return
    
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel Files", "*.xlsx")])
    if output_file:
        export_table_to_excel(database_path.get(), table_var.get(), output_file)

# Initialize GUI application
root = Tk()
root.title("Access to Excel Exporter")

# Variables for storing user selections
database_path = StringVar()
table_var = StringVar()

# GUI Layout
Label(root, text="Step 1: Select Access Database").grid(row=0, column=0, padx=10, pady=5)
Button(root, text="Browse...", command=select_file).grid(row=0, column=1, padx=10)

Label(root, text="Step 2: Select Table").grid(row=1, column=0, padx=10, pady=5)
table_menu = OptionMenu(root, table_var, [])
table_menu.grid(row=1, column=1, padx=10)

Label(root, text="Step 3: Export to Excel").grid(row=2, column=0, padx=10, pady=5)
Button(root, text="Export", command=export_data).grid(row=2, column=1, padx=10)

# Start the GUI event loop
root.mainloop()

