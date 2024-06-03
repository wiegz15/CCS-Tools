import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Function to create the progress window
def show_progress_window(title="Processing", message="Please wait..."):
    progress_window = tk.Toplevel()
    progress_window.title(title)
    progress_window.geometry("300x100")
    progress_window.resizable(False, False)
    progress_label = ttk.Label(progress_window, text=message)
    progress_label.pack(pady=10)
    progress_bar = ttk.Progressbar(progress_window, mode='determinate')
    progress_bar.pack(pady=10, padx=10, fill='x')
    return progress_window, progress_label, progress_bar

# Function to update the progress bar
def update_progress(progress_data, text, percent):
    progress_window, progress_label, progress_bar = progress_data
    progress_label.config(text=text)
    progress_bar['value'] = percent
    progress_window.update()

# Function to execute selected scripts
def execute_scripts(scripts, script_dir):
    progress_data = show_progress_window(title="Executing Scripts", message="Starting script executions...")
    total_scripts = len(scripts)
    for idx, script in enumerate(scripts):
        script_path = os.path.join(script_dir, script)
        update_progress(progress_data, f"Running {script}...", (idx / total_scripts) * 100)
        subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', script_path], check=True)
    update_progress(progress_data, "Scripts execution completed", 100)
    progress_data[0].destroy()
    messagebox.showinfo("Execution Complete", "Scripts execution completed")
    if messagebox.askyesno("Open File?", "Do you want to open the AD_Output.xlsx file now?"):
        excel_path = os.path.join(reports_dir, "AD_Output.xlsx")
        if os.path.exists(excel_path):
            os.startfile(excel_path)
        else:
            messagebox.showerror("File Not Found", "AD_Output.xlsx file not found.")

# Function to create the main window
def create_main_window():
    window = tk.Tk()
    window.title("AD Tools")
    window.geometry("400x860")
    window.resizable(False, False)
    
    tab_control = ttk.Notebook(window)
    
    # Tab 1: AD Health
    tab1 = ttk.Frame(tab_control)
    tab_control.add(tab1, text="AD Health")
    
    tab1_frame = ttk.Frame(tab1)
    tab1_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    select_all_var = tk.BooleanVar()
    
    select_all_cb = ttk.Checkbutton(tab1_frame, text="Select All", variable=select_all_var)
    select_all_cb.pack(anchor='w')
    
    ad_scripts = os.listdir(ad_scripts_dir)
    script_vars = []
    for script in ad_scripts:
        script_var = tk.BooleanVar()
        cb = ttk.Checkbutton(tab1_frame, text=script, variable=script_var)
        cb.pack(anchor='w')
        script_vars.append((script, script_var))
    
    def select_all_changed():
        for _, var in script_vars:
            var.set(select_all_var.get())
    
    select_all_var.trace_add('write', lambda *_: select_all_changed())
    
    execute_button = ttk.Button(tab1_frame, text="Execute", command=lambda: execute_scripts([script for script, var in script_vars if var.get()], ad_scripts_dir))
    execute_button.pack(pady=5)
    
    open_report_button = ttk.Button(tab1_frame, text="Open Report", command=lambda: os.startfile(os.path.join(reports_dir, "AD_Output.xlsx")) if os.path.exists(os.path.join(reports_dir, "AD_Output.xlsx")) else messagebox.showerror("File Not Found", "AD_Output.xlsx file not found."))
    open_report_button.pack(pady=5)
    
    reset_button = ttk.Button(tab1_frame, text="Reset", command=lambda: os.remove(os.path.join(reports_dir, "AD_Output.xlsx")) if os.path.exists(os.path.join(reports_dir, "AD_Output.xlsx")) else messagebox.showinfo("File Not Found", "AD_Output.xlsx file not found."))
    reset_button.pack(pady=5)
    
    # Tab 2: Other Reports
    tab2 = ttk.Frame(tab_control)
    tab_control.add(tab2, text="Other")
    
    tab2_frame = ttk.Frame(tab2)
    tab2_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    select_all_var2 = tk.BooleanVar()
    
    select_all_cb2 = ttk.Checkbutton(tab2_frame, text="Select All", variable=select_all_var2)
    select_all_cb2.pack(anchor='w')
    
    other_scripts = os.listdir(other_scripts_dir)
    script_vars2 = []
    for script in other_scripts:
        script_var = tk.BooleanVar()
        cb = ttk.Checkbutton(tab2_frame, text=script, variable=script_var)
        cb.pack(anchor='w')
        script_vars2.append((script, script_var))
    
    def select_all_changed2():
        for _, var in script_vars2:
            var.set(select_all_var2.get())
    
    select_all_var2.trace_add('write', lambda *_: select_all_changed2())
    
    execute_button2 = ttk.Button(tab2_frame, text="Execute", command=lambda: execute_scripts([script for script, var in script_vars2 if var.get()], other_scripts_dir))
    execute_button2.pack(pady=5)
    
    tab_control.pack(expand=1, fill='both')
    
    exit_button = ttk.Button(window, text="Exit", command=window.quit)
    exit_button.pack(pady=5)
    
    window.mainloop()

# Define paths based on the script location
script_directory = os.path.dirname(os.path.abspath(__file__))
parent_directory = os.path.dirname(script_directory)
reports_dir = os.path.join(parent_directory, "Reports")
ad_scripts_dir = os.path.join(script_directory, "Scripts")
other_scripts_dir = os.path.join(ad_scripts_dir, "Other")

# Ensure the Reports directory exists
os.makedirs(reports_dir, exist_ok=True)
# Ensure the AD Scripts directory exists
os.makedirs(ad_scripts_dir, exist_ok=True)
# Ensure the Other Scripts directory exists
os.makedirs(other_scripts_dir, exist_ok=True)

# Delete the Output.xlsx file if it exists at the start of the script
excel_path = os.path.join(reports_dir, "AD_Output.xlsx")
if os.path.exists(excel_path):
    os.remove(excel_path)

# Run the main window
create_main_window()
