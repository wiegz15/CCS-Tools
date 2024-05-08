# VMware and Active Directory Management Tool

## Overview
This PowerShell script creates a graphical user interface (GUI) for managing VMware and Active Directory tasks. It provides a robust suite of tools for IT administrators to manage virtual environments and directory services efficiently. The GUI simplifies the process of connecting to vCenter, executing scripts, and managing reports.

## Features

### Environment Setup and Validation
- Validates and creates necessary directories for scripts and reports.
- Checks for and potentially creates a configuration file (`tools.ini`) to store and retrieve vCenter server information.

### User Interaction
- Uses Windows Forms and Presentation Framework for a responsive GUI experience.
- Allows users to input the vCenter server name if not previously configured.

### Connectivity
- Connects to vCenter using specified credentials, handling both successful connections and errors.
- Displays connection status dynamically within the GUI.

### Script Execution
- Provides a structured interface to execute multiple scripts related to VMware and Active Directory.
- Includes "Select All" functionality for batch operations.
- Displays progress through a progress bar during script executions.

### Report Handling
- Facilitates opening and resetting report files directly from the GUI.

### Extensibility
- The script is designed to be easily extendable for additional scripts and functionalities.

## Getting Started

### Prerequisites
- PowerShell 5.1 or higher.
- VMware PowerCLI module installed.
- Proper permissions to connect to vCenter and execute scripts.

### Installation
1. Download the script files to your local machine.
2. Ensure all script paths and dependencies are correctly set up in the script.

### Usage
Run the script in PowerShell. The GUI will appear with various controls:
- **Connect to vCenter**: Connect to your vCenter server using credentials.
- **Execute**: Run selected scripts from the available list.
- **Open Report**: View generated reports.
- **Exit**: Properly disconnect and close the application.

