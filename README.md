# Outlook-Export
This Python script automates the process of exporting emails and attachments from Microsoft Outlook into organized directories onto Windows system.

# Script Description
The provided python script is an inspiration from purduevin: "https://github.com/purduevin/OutlookExport". This variation provides an alternative option where the output for emails will be generated to a .MSG format, where users can review the archived emails by clicking each file and open it in Outlook.

# System Requirements and Installation Guide:

Operating System Requirement: Designed specifically for Windows OS due to the dependence on pywin32 and Outlook.

Python Installation:
- Download & Install Python:
	1. Go to the Python Releases for Windows page on Python's official website.
	2. Click on "Download Windows installer".
	
Install pywin32:

    pip install pywin32

Running the Script:
- Ensure Microsoft Outlook is installed and configured with your email account.
- (Optional) Run the script with administrative privileges to enable necessary permissions for accessing Outlook and performing file operations.

This setup ensures that the script can be run efficiently on any compatible Windows machine, providing a robust tool for exporting and archiving email data from Microsoft Outlook.

# Code Usage

Execution method:
    
    python outlook-export-MSG.py -f "FOLDER_NAME" [--single-folder]

Options:


   - -f FOLDER_NAME "Specify which folder from outlook you want to extract from"    
   - --single-folder (Optional function) "Specify if you want all the output within 1 folder. Default setting will separate each email into 1 subfolder based on their timestamp"
