import tkinter as tk
from tkinter import messagebox
import subprocess
import sys
import os
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([script] + sys.argv[1:])
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable.replace('python.exe', 'pythonw.exe'), params, None, 1)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run as admin: {str(e)}")
        sys.exit()

def create_button(frame, text, command, row, column, tooltip=None):
    button = tk.Button(frame, text=text, command=command)
    button.grid(row=row, column=column, padx=10, pady=10, sticky='ew')
    if tooltip:
        create_tooltip(button, tooltip)
    return button

def create_label(frame, text, row, column, color='black'):
    label = tk.Label(frame, text=text, fg=color)
    label.grid(row=row, column=column, padx=10, pady=5, sticky='w')
    return label

def run_script(script_path):
    if not os.path.exists(script_path):
        messagebox.showerror("Error", f"Script not found: {script_path}")
        return
    try:
        result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path], check=True, capture_output=True, text=True)
        messagebox.showinfo("Success", f"Script {script_path} executed successfully.\nOutput:\n{result.stdout}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to run {script_path}\nError:\n{e.stderr}")

def center_popup(popup, root):
    popup.update_idletasks()
    width = popup.winfo_width()
    height = popup.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    popup.geometry(f'{width}x{height}+{x}+{y}')

def check_tools_status():
    popup = tk.Toplevel()
    popup.title("Checking Tool Status")
    label = tk.Label(popup, text="Checking Tool Status. Please wait...")
    label.pack(padx=20, pady=20)
    popup.update()
    center_popup(popup, root)
    popup.attributes("-topmost", True)

    script_path = os.path.join(os.getcwd(), "Update", "ToolsStatus.ps1")
    try:
        result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path], check=True, capture_output=True, text=True)
        status_lines = result.stdout.strip().split('\n')
        for line in status_lines:
            if "Not Installed" in line:
                color = "red"
            elif "Installed" in line:
                color = "green"
            status_label = tk.Label(status_frame, text=line, fg=color)
            status_label.pack(anchor='w')
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to check tools status\nError:\n{e.stderr}")
    finally:
        popup.destroy()

def create_tooltip(widget, text):
    tooltip = tk.Toplevel(widget)
    tooltip.withdraw()
    tooltip.overrideredirect(True)
    label = tk.Label(tooltip, text=text, background="yellow", relief='solid', borderwidth=1, wraplength=180)
    label.pack()

    def enter(event):
        x, y, _, _ = widget.bbox("insert")
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 25
        tooltip.geometry(f"+{x}+{y}")
        tooltip.deiconify()

    def leave(event):
        tooltip.withdraw()

    widget.bind("<Enter>", enter)
    widget.bind("<Leave>", leave)

def main():
    run_as_admin()

    global root
    root = tk.Tk()
    root.title('CCS Tools Launcher')
    root.geometry('420x450')

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10, fill='both', expand=True)

    global status_frame
    status_container = tk.Frame(root)
    status_container.pack(padx=10, pady=10, fill='both', expand=True)

    create_label(status_container, 'Tools Status:', 0, 0, color='blue')
    status_frame = tk.Frame(status_container)
    status_frame.grid(row=1, column=0, sticky='nsew')

    vmware_path = os.path.join(os.getcwd(), "Vmware")
    ad_path = os.path.join(os.getcwd(), "AD")
    update_path = os.path.join(os.getcwd(), "Update")
    rsat_path = os.path.join(os.getcwd(), "Update")

    create_button(frame, 'VMware Toolset', lambda: run_script(os.path.join(vmware_path, "vmware_launcher_new.ps1")), 5, 0, "Launch VMware Toolset")
    create_button(frame, 'AD Toolset', lambda: run_script(os.path.join(ad_path, "AD_launcher.ps1")), 5, 1, "Launch Active Directory Toolset")
    create_button(frame, 'Check PS Modules', lambda: run_script(os.path.join(update_path, "UpdateModule.ps1")), 7, 0, "Check and Update PowerShell Modules")
    create_button(frame, 'Install RSAT Tools', lambda: run_script(os.path.join(rsat_path, "RSAT Tools.ps1")), 7, 1, "Install RSAT Tools")

    check_tools_status()

    root.mainloop()

if __name__ == "__main__":
    main()
