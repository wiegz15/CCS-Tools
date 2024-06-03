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
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run as admin: {str(e)}")
        sys.exit()

def create_button(frame, text, command, row, column):
    button = tk.Button(frame, text=text, command=command)
    button.grid(row=row, column=column, padx=10, pady=10, sticky='ew')
    return button

def create_label(frame, text, row, column, color='black'):
    label = tk.Label(frame, text=text, fg=color)
    label.grid(row=row, column=column, padx=10, pady=5, sticky='w')
    return label

def install_module(module_name):
    try:
        subprocess.run(["powershell", "-Command", f"Install-Module -Name {module_name} -Force"], check=True)
        messagebox.showinfo("Info", f"{module_name} installed successfully")
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", f"Failed to install {module_name}")

def check_module_installed(module_name):
    try:
        result = subprocess.run(["powershell", "-Command", f"Get-Module -ListAvailable -Name {module_name}"], capture_output=True, text=True)
        return module_name in result.stdout
    except subprocess.CalledProcessError:
        return False

def check_rsat_server_installed():
    rsat_server_features = [
        "RSAT",
        "DNS-Server-Tools",
        "RSAT-AD-Tools-Feature",
        "RSAT-ADDS-Tools-Feature",
        "DirectoryServices-DomainController-Tools",
        "DirectoryServices-ADAM-Tools",
        "RSAT-DHCP"
    ]
    for feature in rsat_server_features:
        result = subprocess.run(["dism", "/online", "/get-featureinfo", f"/featurename:{feature}"], capture_output=True, text=True)
        if "State : Enabled" not in result.stdout:
            return False
    return True

def check_rsat_win_installed():
    rsat_win_features = [
        "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0",
        "Rsat.BitLocker.Recovery.Tools~~~~0.0.1.0",
        "Rsat.CertificateServices.Tools~~~~0.0.1.0",
        "Rsat.DHCP.Tools~~~~0.0.1.0",
        "Rsat.Dns.Tools~~~~0.0.1.0",
        "Rsat.FailoverCluster.Management.Tools~~~~0.0.1.0",
        "Rsat.FileServices.Tools~~~~0.0.1.0",
        "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0",
        "Rsat.IPAM.Client.Tools~~~~0.0.1.0",
        "Rsat.LLDP.Tools~~~~0.0.1.0",
        "Rsat.NetworkController.Tools~~~~0.0.1.0",
        "Rsat.NetworkLoadBalancing.Tools~~~~0.0.1.0",
        "Rsat.Print.Management.Tools~~~~0.0.1.0",
        "Rsat.RemoteAccess.Management.Tools~~~~0.0.1.0",
        "Rsat.RemoteDesktop.Services.Tools~~~~0.0.1.0",
        "Rsat.ServerManager.Tools~~~~0.0.1.0",
        "Rsat.Shielded.VM.Tools~~~~0.0.1.0",
        "Rsat.StorageMigrationService.Management.Tools~~~~0.0.1.0",
        "Rsat.StorageReplica.Tools~~~~0.0.1.0",
        "Rsat.SystemInsights.Tools~~~~0.0.1.0",
        "Rsat.VolumeActivation.Tools~~~~0.0.1.0",
        "Rsat.WSUS.Tools~~~~0.0.1.0"
    ]
    for feature in rsat_win_features:
        result = subprocess.run(["powershell", "-Command", f"(Get-WindowsCapability -Online -Name {feature}).State"], capture_output=True, text=True)
        if "Installed" not in result.stdout:
            return False
    return True

def run_script(script_path):
    try:
        subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path], check=True)
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", f"Failed to run {script_path}")

def install_rsat_server():
    try:
        subprocess.run([
            "dism", "/online", "/enable-feature",
            "/featurename:RSAT",
            "/featurename:DNS-Server-Tools",
            "/featurename:RSAT-AD-Tools-Feature",
            "/featurename:RSAT-ADDS-Tools-Feature",
            "/featurename:DirectoryServices-DomainController-Tools",
            "/featurename:DirectoryServices-ADAM-Tools",
            "/featurename:RSAT-DHCP",
            "/all"
        ], check=True)
        messagebox.showinfo("Info", "RSAT features installed successfully")
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "Failed to install RSAT features")

def install_rsat_win():
    rsat_features = [
        "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0",
        "Rsat.BitLocker.Recovery.Tools~~~~0.0.1.0",
        "Rsat.CertificateServices.Tools~~~~0.0.1.0",
        "Rsat.DHCP.Tools~~~~0.0.1.0",
        "Rsat.Dns.Tools~~~~0.0.1.0",
        "Rsat.FailoverCluster.Management.Tools~~~~0.0.1.0",
        "Rsat.FileServices.Tools~~~~0.0.1.0",
        "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0",
        "Rsat.IPAM.Client.Tools~~~~0.0.1.0",
        "Rsat.LLDP.Tools~~~~0.0.1.0",
        "Rsat.NetworkController.Tools~~~~0.0.1.0",
        "Rsat.NetworkLoadBalancing.Tools~~~~0.0.1.0",
        "Rsat.Print.Management.Tools~~~~0.0.1.0",
        "Rsat.RemoteAccess.Management.Tools~~~~0.0.1.0",
        "Rsat.RemoteDesktop.Services.Tools~~~~0.0.1.0",
        "Rsat.ServerManager.Tools~~~~0.0.1.0",
        "Rsat.Shielded.VM.Tools~~~~0.0.1.0",
        "Rsat.StorageMigrationService.Management.Tools~~~~0.0.1.0",
        "Rsat.StorageReplica.Tools~~~~0.0.1.0",
        "Rsat.SystemInsights.Tools~~~~0.0.1.0",
        "Rsat.VolumeActivation.Tools~~~~0.0.1.0",
        "Rsat.WSUS.Tools~~~~0.0.1.0"
    ]
    try:
        for feature in rsat_features:
            subprocess.run(["dism", "/online", "/add-capability", f"Name={feature}"], check=True)
        messagebox.showinfo("Info", "RSAT features installed successfully")
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "Failed to install RSAT features")

def main():
    run_as_admin()

    root = tk.Tk()
    root.title('CCS Tools Launcher')
    root.geometry('420x450')

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10, fill='both', expand=True)

    vmware_path = os.path.join(os.getcwd(), "Vmware")
    ad_path = os.path.join(os.getcwd(), "AD")

    create_label(frame, f"ImportExcel: {'Installed' if check_module_installed('ImportExcel') else 'Not Installed'}", 0, 0)
    create_label(frame, f"PowerCLI: {'Installed' if check_module_installed('VMware.PowerCLI') else 'Not Installed'}", 0, 1)

    create_button(frame, 'Install ImportExcel', lambda: install_module('ImportExcel'), 1, 0)
    create_button(frame, 'Install PowerCLI', lambda: install_module('VMware.PowerCLI'), 1, 1)

    create_label(frame, f"RSAT (Server): {'Installed' if check_rsat_server_installed() else 'Not Installed'}", 2, 0)
    create_button(frame, 'Install Features for Servers Only', install_rsat_server, 3, 0)

    create_label(frame, f"RSAT (Win 10/11): {'Installed' if check_rsat_win_installed() else 'Not Installed'}", 2, 1)
    create_button(frame, 'Install RSAT for Windows 10/11', install_rsat_win, 3, 1)

    if check_module_installed('ImportExcel') and check_module_installed('VMware.PowerCLI') and (check_rsat_server_installed() or check_rsat_win_installed()):
        create_label(frame, "✔ Pre-reqs installed", 4, 0, 'green')

    create_button(frame, 'VMware Toolset', lambda: run_script(os.path.join(vmware_path, "vmware_launcher_new.ps1")), 5, 0)
    create_button(frame, 'AD Toolset', lambda: run_script(os.path.join(ad_path, "AD_launcher.ps1")), 5, 1)

    root.mainloop()

if __name__ == "__main__":
    main()
