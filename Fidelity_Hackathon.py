import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from pandastable import Table
from tkinter import simpledialog

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV File Operations")
        self.file_path = None
        self.df = None
        
        self.load_button = tk.Button(root, text="Load Excel/CSV File", command=self.load_file)
        self.load_button.pack(pady=10)
        
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill=tk.BOTH, expand=1)
        
        self.filter_button = tk.Button(root, text="Filter and Select", command=self.filter_and_select)
        self.filter_button.pack(pady=10)
        
        self.date_format_button = tk.Button(root, text="Change Date Format", command=self.change_date_format)
        self.date_format_button.pack(pady=10)
        
        self.save_button = tk.Button(root, text="Save Changes", command=self.save_changes)
        self.save_button.pack(pady=10)

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if self.file_path:
            if self.file_path.endswith('.csv'):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)
            self.show_table(self.df)
    
    def show_table(self, df):
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        table = Table(self.table_frame, dataframe=df, showtoolbar=True, showstatusbar=True)
        table.show()

    def filter_and_select(self):
        if self.df is not None:
            columns = self.df.columns.tolist()
            selected_columns = simpledialog.askstring("Filter Columns", f"Enter column names to keep (comma separated):\nAvailable columns: {columns}")
            if selected_columns:
                selected_columns = [col.strip() for col in selected_columns.split(',')]
                filtered_df = self.df[selected_columns]
                self.show_table(filtered_df)
                save_option = messagebox.askyesno("Save Filtered Data", "Do you want to save the filtered data to a new file?")
                if save_option:
                    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
                    if save_path:
                        if save_path.endswith('.csv'):
                            filtered_df.to_csv(save_path, index=False)
                        else:
                            filtered_df.to_excel(save_path, index=False)
        else:
            messagebox.showerror("Error", "No file loaded.")

    def change_date_format(self):
        if self.df is not None:
            date_columns = self.df.select_dtypes(include=['datetime', 'datetime64']).columns.tolist()
            if date_columns:
                for col in date_columns:
                    new_format = simpledialog.askstring("Change Date Format", f"Enter new date format for column {col} (e.g., '%Y-%m-%d'):")
                    if new_format:
                        self.df[col] = pd.to_datetime(self.df[col]).dt.strftime(new_format)
                self.show_table(self.df)
            else:
                messagebox.showinfo("No Date Columns", "No date columns found in the loaded file.")
        else:
            messagebox.showerror("Error", "No file loaded.")

    def save_changes(self):
        if self.df is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
            if save_path:
                if save_path.endswith('.csv'):
                    self.df.to_csv(save_path, index=False)
                else:
                    self.df.to_excel(save_path, index=False)
        else:
            messagebox.showerror("Error", "No file loaded.")

root = tk.Tk()
app = ExcelApp(root)
root.mainloop()
