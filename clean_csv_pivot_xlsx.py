import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import numpy as np
from openpyxl import load_workbook

class PivotTableApp:
  def __init__(self, master):
      self.master = master
      self.master.title("Pivot Table Creator")
      self.master.geometry("1000x600")

      self.style = ttk.Style()
      self.style.theme_use('clam')

      self.df = None
      self.pivot_table = None
      self.filters = []

      self.create_widgets()

  def create_widgets(self):
      main_frame = ttk.Frame(self.master, padding="10")
      main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
      self.master.columnconfigure(0, weight=1)
      self.master.rowconfigure(0, weight=1)

      # File selection and skip rows
      file_frame = ttk.Frame(main_frame, padding="5")
      file_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
      ttk.Label(file_frame, text="Skip Rows:").pack(side=tk.LEFT)
      self.skip_rows_var = tk.StringVar(value="0")
      ttk.Entry(file_frame, textvariable=self.skip_rows_var, width=5).pack(side=tk.LEFT, padx=5)
      self.file_button = ttk.Button(file_frame, text="Select File", command=self.load_file)
      self.file_button.pack(side=tk.LEFT, padx=5)

      # Filters
      for i in range(3):
          filter_frame = ttk.Frame(main_frame, padding="5")
          filter_frame.grid(row=1, column=i, sticky=(tk.W, tk.E, tk.N, tk.S))
          ttk.Label(filter_frame, text=f"Filter {i+1}:").grid(row=0, column=0, sticky=tk.W)
          column_dropdown = ttk.Combobox(filter_frame)
          column_dropdown.grid(row=1, column=0, sticky=(tk.W, tk.E))
          values_listbox = tk.Listbox(filter_frame, selectmode=tk.MULTIPLE, exportselection=0)
          values_listbox.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
          scrollbar = ttk.Scrollbar(filter_frame, orient="vertical", command=values_listbox.yview)
          scrollbar.grid(row=2, column=1, sticky=(tk.N, tk.S))
          values_listbox.config(yscrollcommand=scrollbar.set)
          select_all_var = tk.BooleanVar()
          select_all_check = ttk.Checkbutton(filter_frame, text="Select All", variable=select_all_var, command=lambda lb=values_listbox, var=select_all_var: self.toggle_select_all(lb, var))
          select_all_check.grid(row=3, column=0, sticky=tk.W)
          filter_frame.columnconfigure(0, weight=1)
          filter_frame.rowconfigure(2, weight=1)
          self.filters.append((column_dropdown, values_listbox, select_all_var))
          column_dropdown.bind("<<ComboboxSelected>>", lambda e, i=i: self.update_filter_values(i))

      # Pivot table options
      pivot_frame = ttk.Frame(main_frame, padding="5")
      pivot_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E))
      
      ttk.Label(pivot_frame, text="Index:").grid(row=0, column=0)
      self.index_listbox = tk.Listbox(pivot_frame, selectmode=tk.MULTIPLE, exportselection=0, height=4)
      self.index_listbox.grid(row=0, column=1, sticky=(tk.W, tk.E))
      index_scrollbar = ttk.Scrollbar(pivot_frame, orient="vertical", command=self.index_listbox.yview)
      index_scrollbar.grid(row=0, column=2, sticky=(tk.N, tk.S))
      self.index_listbox.config(yscrollcommand=index_scrollbar.set)

      ttk.Label(pivot_frame, text="Columns:").grid(row=0, column=3)
      self.columns_listbox = tk.Listbox(pivot_frame, selectmode=tk.MULTIPLE, exportselection=0, height=4)
      self.columns_listbox.grid(row=0, column=4, sticky=(tk.W, tk.E))
      columns_scrollbar = ttk.Scrollbar(pivot_frame, orient="vertical", command=self.columns_listbox.yview)
      columns_scrollbar.grid(row=0, column=5, sticky=(tk.N, tk.S))
      self.columns_listbox.config(yscrollcommand=columns_scrollbar.set)

      ttk.Label(pivot_frame, text="Values:").grid(row=0, column=6)
      self.values_listbox = tk.Listbox(pivot_frame, selectmode=tk.MULTIPLE, exportselection=0, height=4)
      self.values_listbox.grid(row=0, column=7, sticky=(tk.W, tk.E))
      values_scrollbar = ttk.Scrollbar(pivot_frame, orient="vertical", command=self.values_listbox.yview)
      values_scrollbar.grid(row=0, column=8, sticky=(tk.N, tk.S))
      self.values_listbox.config(yscrollcommand=values_scrollbar.set)

      ttk.Label(pivot_frame, text="Aggregate Function:").grid(row=1, column=0, columnspan=2)
      self.aggfunc_var = tk.StringVar()
      self.aggfunc_dropdown = ttk.Combobox(pivot_frame, textvariable=self.aggfunc_var, values=['sum', 'mean', 'count', 'max', 'min'])
      self.aggfunc_dropdown.grid(row=1, column=2, columnspan=2)

      # Export button
      self.export_button = ttk.Button(main_frame, text="Export to Excel", command=self.export_to_excel)
      self.export_button.grid(row=3, column=0, columnspan=3, pady=10)

      for child in main_frame.winfo_children():
          child.grid_configure(padx=5, pady=5)

  def toggle_select_all(self, listbox, var):
      if var.get():
          listbox.select_set(0, tk.END)
      else:
          listbox.selection_clear(0, tk.END)

  def load_file(self):
      file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
      if file_path:
          try:
              skip_rows = int(self.skip_rows_var.get())
              if file_path.endswith('.csv'):
                  self.df = pd.read_csv(file_path, skiprows=skip_rows)
              else:
                  self.df = pd.read_excel(file_path, skiprows=skip_rows)
              
              self.columns = list(self.df.columns)
              self.update_dropdowns()
              messagebox.showinfo("Success", "File loaded successfully!")
          except Exception as e:
              messagebox.showerror("Error", f"Failed to load file: {str(e)}")

  def update_dropdowns(self):
      for column_dropdown, _, _ in self.filters:
          column_dropdown['values'] = self.columns
      for listbox in [self.index_listbox, self.columns_listbox, self.values_listbox]:
          listbox.delete(0, tk.END)
          for column in self.columns:
              listbox.insert(tk.END, column)

  def update_filter_values(self, filter_index):
      column_dropdown, values_listbox, select_all_var = self.filters[filter_index]
      column = column_dropdown.get()
      if column:
          values = self.df[column].unique().tolist()
          values_listbox.delete(0, tk.END)
          for value in values:
              values_listbox.insert(tk.END, value)
          select_all_var.set(False)

  def get_selected_items(self, listbox):
      return [listbox.get(i) for i in listbox.curselection()]

  def apply_filters(self):
      filtered_df = self.df.copy()
      for column_dropdown, values_listbox, _ in self.filters:
          column = column_dropdown.get()
          selected_values = self.get_selected_items(values_listbox)
          if column and selected_values:
              filtered_df = filtered_df[filtered_df[column].isin(selected_values)]
      return filtered_df

  def create_pivot_table(self):
      if self.df is not None:
          filtered_df = self.apply_filters()
          index = self.get_selected_items(self.index_listbox)
          columns = self.get_selected_items(self.columns_listbox)
          values = self.get_selected_items(self.values_listbox)
          aggfunc = self.aggfunc_var.get()

          if not index or not values:
              messagebox.showerror("Error", "Please select at least one Index and one Value for the pivot table.")
              return False

          try:
              self.pivot_table = pd.pivot_table(filtered_df, values=values, index=index, columns=columns, aggfunc=aggfunc)
              self.pivot_table = self.pivot_table.replace(np.nan, '', regex=True)  # Replace NaN with empty string
              return True
          except Exception as e:
              messagebox.showerror("Error", f"Failed to create pivot table: {str(e)}")
              return False

  def export_to_excel(self):
      if self.create_pivot_table():
          file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
          if file_path:
              try:
                  # Convert pivot table to Excel
                  self.pivot_table.to_excel(file_path)
                  
                  # Open the saved file and adjust column widths
                  wb = load_workbook(file_path)
                  ws = wb.active
                  
                  column_width_adjusted = True
                  try:
                      for column in ws.columns:
                          max_length = 0
                          column_letter = column[0].column_letter
                          for cell in column:
                              try:
                                  if len(str(cell.value)) > max_length:
                                      max_length = len(cell.value)
                              except:
                                  pass
                          adjusted_width = (max_length + 2)
                          ws.column_dimensions[column_letter].width = adjusted_width
                      
                      wb.save(file_path)
                  except AttributeError as e:
                      print(f"Warning: Could not adjust column widths. Error: {str(e)}")
                      column_width_adjusted = False
                  
                  if column_width_adjusted:
                      messagebox.showinfo("Success", f"Exported pivot table to: {file_path}")
                  else:
                      messagebox.showinfo("Success with Warning", f"Exported pivot table to: {file_path}\n\nNote: Column widths could not be adjusted automatically.")
                  print(f"Exported pivot table to: {file_path}")
              except Exception as e:
                  messagebox.showerror("Error", f"Failed to export pivot table: {str(e)}")

if __name__ == "__main__":
  root = tk.Tk()
  app = PivotTableApp(root)
  root.mainloop()

# Created/Modified files during execution:
# The script will create an Excel file when exporting the pivot table.
# The file name will be determined by the user during the save dialog.
