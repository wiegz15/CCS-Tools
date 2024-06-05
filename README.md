# CCS Tools Launcher

CCS Tools Launcher is a Python-based application that provides a graphical user interface (GUI) for running various PowerShell scripts related to VMware, Active Directory, and other tools. It ensures that the required scripts are executed with administrative privileges and checks the status of installed tools.

## Features

- **Run as Admin:** Ensures the application runs with administrative privileges.
- **Extract Required Folders:** Extracts the necessary folders for running scripts.
- **Run PowerShell Scripts:** Provides buttons to run specific PowerShell scripts for different toolsets.
- **Check Tools Status:** Displays the status of installed tools and modules.
- **Tooltips:** Provides tooltips for buttons to explain their functionality.

## Requirements

- Python 3.x
- Tkinter library (usually included with Python)
- PowerShell
- Required PowerShell scripts in `Update`, `AD`, and `Vmware` folders.

## GUI Components

### Buttons

- **VMware Toolset:** Launches the VMware Toolset by running the `vmware_launcher_new.ps1` script.
- **AD Toolset:** Launches the Active Directory Toolset by running the `AD_launcher.ps1` script.
- **Update Tools:** Checks and updates tools by running the `Update Tools.ps1` script.
- **Install PS Modules:** Checks and updates PowerShell modules by running the `Install Modules.ps1` script.
- **Install RSAT Tools:** Installs RSAT tools by running the `Install RSAT Tools.ps1` script.

### Tools Status

Displays the status of installed tools, indicating whether they are installed or not.

## Code Overview

### Main Functions

- **is_admin:** Checks if the application is running with administrative privileges.
- **run_as_admin:** Restarts the application with administrative privileges if not already running as admin.
- **extract_folder:** Extracts the specified folder to the current working directory.
- **create_button:** Creates a button in the GUI.
- **create_label:** Creates a label in the GUI.
- **run_script:** Runs the specified PowerShell script and displays the output or error.
- **center_popup:** Centers a popup window on the screen.
- **check_tools_status:** Checks the status of installed tools by running the `ToolsStatus.ps1` script.
- **create_tooltip:** Creates a tooltip for a widget.
