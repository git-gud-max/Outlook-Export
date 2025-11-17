# Outlook-Export
This Python script automates the process of exporting emails and attachments from Microsoft Outlook into organized directories onto Windows system.

# Script Description
The provided python script is an inspiration from purduevin: "https://github.com/purduevin/OutlookExport". This variation provides an alternative option where the output for emails will be generated to a .MSG format, where users can review the archived emails by clicking each file and open it in Outlook.

# System Requirements and Installation Guide:

Operating System Requirement: Designed specifically for Windows OS due to the dependence on pywin32 and Outlook.

Python Installation:
Download Python:
    Go to the Python Releases for Windows page on Python's official website.
    Click on "Download Windows installer".
Install Python:
    Run the downloaded installer.
    Make sure to check "Add Python 3.x to PATH" at the bottom of the installation window to automatically add Python to your environment variables.
    Click "Install Now".

Adding Python and pip to PATH Manually: If you didn’t add Python to your PATH during the installation, you can add it manually:

Open the Start Search, type env, and select "Edit the system environment variables" or "Edit environment variables for your account".
Under System Properties, click on the "Environment Variables…" button.
Find the 'Path' variable in the "System variables" section and click "Edit…".
Add Python path:
    Click "New" and add the path to the folder where Python is installed, e.g., C:\Users\<Username>\AppData\Local\Programs\Python\Python39.
    Add another new line for the Scripts directory, e.g., C:\Users\<Username>\AppData\Local\Programs\Python\Python39\Scripts.
    Click OK on all dialogs to close them.

Install pywin32:

    pip install pywin32

Running the Script:
    Ensure Microsoft Outlook is installed and configured with your email account.
    Run the script with administrative privileges to enable necessary permissions for accessing Outlook and performing file operations.

This setup ensures that the script can be run efficiently on any compatible Windows machine, providing a robust tool for exporting and archiving email data from Microsoft Outlook.

# Code Usage

Execution method:
    
    python outlook-export-MSG.py -f "FOLDER_NAME" [--single-folder]

Options:
    -f FOLDER_NAME "Specify which folder from outlook you want to extract from"    
    --single-folder (Optional function) "Specify if you want all the output within 1 folder. Default setting will separate each email into 1 subfolder based on their timestamp"
