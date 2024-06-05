# CCS Tools Launcher

CCS Tools Launcher is a Python-based application that provides a graphical user interface (GUI) for running various PowerShell scripts related to VMware, Active Directory, and other tools. It ensures that the required scripts are executed with administrative privileges and checks the status of installed tools.

## Features

- **Run as Admin:** Ensures the application runs with administrative privileges.
- **Extract Required Folders:** Extracts the necessary folders for running scripts.
- **Run PowerShell Scripts:** Provides buttons to run specific PowerShell scripts for different toolsets.
- **Check Tools Status:** Displays the status of installed tools and modules.
- **Tooltips:** Provides tooltips for buttons to explain their functionality.

## Requirements

- Domain Administrator rights.
- VCenter Admin rights.
- PowerShell
- You may need to fix powershell Execution Policy for now if it fails to run the Powershell Scripts.
- Required PowerShell scripts in `Update`, `AD`, and `Vmware` folders. (These will update with the 'Update Tools' button.)
  -NOTE: The Powershell Install-Module part can take some time depending on the device.

## Usage

When you run the CCS Tools Launcher EXE, it will automatically create the necessary folders it needs to run.

### Steps:

1. **Run the EXE:**

   - Launch the CCS Tools Launcher executable file. It is recommended you run this in its own folder and not right on the Desktop or in Documents.

2. **Update Tools:**

   - Click the **Update Tools** button in the GUI.
   - This will download a portable version of Git and check the repository for updated scripts in the `AD` and `Vmware` folders. The Portable Git download only needs to happen the first time.

3. **Using the Toolsets:**
   - **VMware Toolset:** Click the **VMware Toolset** button to run the `vmware_launcher_new.ps1` script.
   - **AD Toolset:** Click the **AD Toolset** button to run the `AD_launcher.ps1` script.
   - **Install PS Modules:** Click the **Install PS Modules** button to check and update PowerShell modules by running the `Install Modules.ps1` script.
   - **Install RSAT Tools:** Click the **Install RSAT Tools** button to install RSAT tools by running the `Install RSAT Tools.ps1` script.

### Tools Status:

The Tools Status section displays the status of installed tools, indicating whether they are installed or not.

By following these steps, you can ensure that your toolsets are up-to-date and ready to use.

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
