import datetime
from pathlib import Path
import re
import win32com.client
from win32com.client import gencache
import argparse

def create_directory(base_path, folder_name):
    """ Helper function to create a directory for storing email exports. """
    try:
        new_dir = (base_path / folder_name)
        new_dir.mkdir(parents=True, exist_ok=True)
        return new_dir
    except Exception as e:
        print(f"Error creating directory {folder_name}: {str(e)}")
        return None

def save_attachments(attachments, folder_path):
    """ Saves all email attachments to a specified directory. """
    for attachment in attachments:
        try:
            attachment_name = re.sub(r'[^\w\s.]+', '', attachment.FileName)
            attachment.SaveAsFile(str(folder_path / attachment_name))
        except Exception as e:
            print(f"Failed to save attachment {attachment.FileName}: {str(e)}")

def export_emails(folder, output_dir):
    """ Processes and exports emails from a given Outlook folder. """
    messages = folder.Items
    if folder.DefaultItemType == win32com.client.constants.olMailItem:
        try:
            messages.Sort("[ReceivedTime]", True)
        except Exception as e:
            print(f"Could not sort items in folder {folder.Name}: {str(e)}")

    processed_count = 0
    errors_and_skips = []    

    for message in messages:
        email_time = "Unknown"
        try:
            if not hasattr(message, 'ReceivedTime'):
                errors_and_skips.append("Skipped non-mail item with unknown received time.")
                continue
            if message.Class != win32com.client.constants.olMail:
                errors_and_skips.append(f"Skipped non-mail item at {message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}")
                continue

            email_time = message.ReceivedTime.strftime("%Y-%m-%d_%H-%M-%S")
            folder_name = f"{email_time}"
            email_folder = create_directory(output_dir, folder_name)

            if email_folder:
                subject = getattr(message, 'Subject', 'No Subject')
                subject_sanitized = re.sub(r'[^\w\s-]+', '', subject)  # Sanitize subject for filename
                file_name = f"{subject_sanitized}_{email_time}.msg"
                
                # Save the email as a .msg file
                message.SaveAs(str(email_folder / file_name), win32com.client.constants.olMSG)
                
                if message.Attachments.Count > 0:
                    save_attachments(message.Attachments, email_folder)
            processed_count += 1
        except Exception as e:
            errors_and_skips.append(f"Error with item at {email_time}: {str(e)}")

    print(f"Successfully processed {processed_count} emails in folder {folder.Name}.")
    return processed_count, errors_and_skips

def export_emails_to_single_folder(folder, output_dir):
    """ Processes and exports all emails from a given Outlook folder into one directory. """
    messages = folder.Items
    if folder.DefaultItemType == win32com.client.constants.olMailItem:
        try:
            messages.Sort("[ReceivedTime]", True)
        except Exception as e:
            print(f"Could not sort items in folder {folder.Name}: {str(e)}")

    processed_count = 0
    errors_and_skips = []

    # Create a base directory for all emails within the folder
    email_folder = create_directory(output_dir, folder.Name)
    # Create a sub directory for all-in-one emails
    test_base_dir = create_directory(email_folder, "All-In-One")
    att_base_dir = create_directory(test_base_dir, "Attachments")

    for message in messages:
        email_time = "Unknown"
        try:
            if not hasattr(message, 'ReceivedTime'):
                errors_and_skips.append("Skipped non-mail item with unknown received time.")
                continue
            if message.Class != win32com.client.constants.olMail:
                errors_and_skips.append(f"Skipped non-mail item at {message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}")
                continue

            # Sanitize the subject for use in a filename
            subject = getattr(message, 'Subject', 'No Subject')
            subject_sanitized = re.sub(r'[^\w\s-]+', '', subject)
            email_time = message.ReceivedTime.strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"{subject_sanitized}_{email_time}.msg"

            # Save the email as a .msg file within the directory for the entire folder
            message.SaveAs(str(test_base_dir / file_name), win32com.client.constants.olMSG)
            if message.Attachments.Count > 0:
                    save_attachments(message.Attachments, att_base_dir)
            processed_count += 1
        except Exception as e:
            errors_and_skips.append(f"Error with item at {email_time}: {str(e)}")

    print(f"Successfully processed {processed_count} emails in folder {folder.Name} into single directory.")
    return processed_count, errors_and_skips

def main():
    """ Main function to initialize Outlook access and process specified folders. """
    print("Starting the script...")
    
    # Set up command-line argument parsing
    parser = argparse.ArgumentParser(description="Process specified Outlook folders.")
    parser.add_argument('-f', '--folders', nargs='+', help='List of folder names to process', required=True)
    parser.add_argument('--single-folder', action='store_true', help='Export all emails into a single folder')
    args = parser.parse_args()

    specified_folders = set(args.folders)  # Get folder names from command line arguments

    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    base_dir = Path.cwd() / "EmailExports"
    base_dir.mkdir(parents=True, exist_ok=True)

    total_processed = 0
    all_errors_and_skips = []

    all_folders = namespace.Folders
    for store in all_folders:
        print(f"Processing store: {store.Name}")
        store_dir = create_directory(base_dir, store.Name)
        
        folders = store.Folders

        for folder in folders:
            if folder.Name in specified_folders:
                print(f"Processing folder: {folder.Name}")

                if args.single_folder:
                    processed_count, errors_and_skips = export_emails_to_single_folder(folder, store_dir)
                else:
                    folder_dir = create_directory(store_dir, folder.Name)
                    processed_count, errors_and_skips = export_emails(folder, folder_dir)

                total_processed += processed_count
                all_errors_and_skips.extend(errors_and_skips)

    if all_errors_and_skips:
        print("Summary of errors and skipped emails:")
        for error in all_errors_and_skips:
            print(error)
    else:
        print("No errors or skipped emails.")

    print(f"Total emails processed: {total_processed}")
    print("Email export completed.")

if __name__ == "__main__":
    main()
