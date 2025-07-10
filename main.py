#!/usr/bin/env python3
"""
Email Attachment Downloader
This script downloads attachments from email and provides search configuration options
"""

import os
import re
from email_downloader import EmailAttachmentDownloader
from email_config import EMAIL_SETTINGS, DEFAULT_CONFIG
from dotenv import load_dotenv

def display_current_config():
    """Display current search range configuration"""
    print("=== Current Email Configuration ===")
    
    # Show folder
    print(f"Target Folder: {EMAIL_SETTINGS['folder_name']}")
    
    # Show search range settings
    search_range = EMAIL_SETTINGS["search_range"]
    print(f"Search Range Enabled: {search_range['enabled']}")
    print(f"Search Type: {search_range['type']}")
    print(f"Batch Size: {search_range['batch_size']}")
    
    if search_range['type'] == 'date_range':
        date_range = search_range['date_range']
        print(f"Date Range: {date_range['start_date']} to {date_range['end_date']}")
    elif search_range['type'] == 'count_limit':
        print(f"Count Limit: {search_range['count_limit']} emails")
    elif search_range['type'] == 'recent_days':
        print(f"Recent Days: {search_range['recent_days']} days")
    
    print()

def configure_search_range():
    """Interactive configuration of search range"""
    print("=== Email Search Range Configuration ===")
    print("Choose your search range option:")
    print("1. All emails (default)")
    print("2. Date range")
    print("3. Count limit (process only N emails)")
    print("4. Recent days (last N days)")
    print("5. Disable search range (use all emails)")
    
    choice = input("\nEnter your choice (1-5): ").strip()
    
    if choice == "1":
        # All emails
        EMAIL_SETTINGS["search_range"]["enabled"] = False
        EMAIL_SETTINGS["search_range"]["type"] = "all"
        print("✓ Configured to download all emails")
        
    elif choice == "2":
        # Date range
        EMAIL_SETTINGS["search_range"]["enabled"] = True
        EMAIL_SETTINGS["search_range"]["type"] = "date_range"
        
        print("\nEnter date range (format: YYYYMMDD)")
        print("Example: 20240101, 20240315, 20241231")
        
        while True:
            start_date = input("Start date: ").strip()
            end_date = input("End date: ").strip()
            
            # Basic validation for date format
            date_pattern = r'^\d{8}$'
            
            if re.match(date_pattern, start_date) and re.match(date_pattern, end_date):
                EMAIL_SETTINGS["search_range"]["date_range"]["start_date"] = start_date
                EMAIL_SETTINGS["search_range"]["date_range"]["end_date"] = end_date
                print(f"✓ Configured date range: {start_date} to {end_date}")
                break
            else:
                print("❌ Invalid date format. Please use YYYYMMDD format (e.g., 20240101)")
                retry = input("Try again? (y/n): ").lower().strip()
                if retry != 'y':
                    print("Using default date range: 20240101 to 20241231")
                    EMAIL_SETTINGS["search_range"]["date_range"]["start_date"] = "20240101"
                    EMAIL_SETTINGS["search_range"]["date_range"]["end_date"] = "20241231"
                    break
        
    elif choice == "3":
        # Count limit
        EMAIL_SETTINGS["search_range"]["enabled"] = True
        EMAIL_SETTINGS["search_range"]["type"] = "count_limit"
        
        try:
            count_limit = int(input("Enter maximum number of emails to process: ").strip())
            EMAIL_SETTINGS["search_range"]["count_limit"] = count_limit
            print(f"✓ Configured to process maximum {count_limit} emails")
        except ValueError:
            print("Invalid number. Using default of 100 emails.")
            EMAIL_SETTINGS["search_range"]["count_limit"] = 100
            
    elif choice == "4":
        # Recent days
        EMAIL_SETTINGS["search_range"]["enabled"] = True
        EMAIL_SETTINGS["search_range"]["type"] = "recent_days"
        
        try:
            recent_days = int(input("Enter number of recent days to search: ").strip())
            EMAIL_SETTINGS["search_range"]["recent_days"] = recent_days
            print(f"✓ Configured to search last {recent_days} days")
        except ValueError:
            print("Invalid number. Using default of 30 days.")
            EMAIL_SETTINGS["search_range"]["recent_days"] = 30
            
    elif choice == "5":
        # Disable search range
        EMAIL_SETTINGS["search_range"]["enabled"] = False
        print("✓ Disabled search range - will download all emails")
        
    else:
        print("Invalid choice. Using default (all emails).")
        EMAIL_SETTINGS["search_range"]["enabled"] = False
        EMAIL_SETTINGS["search_range"]["type"] = "all"
    
    # Configure batch size only for options that use it
    if choice in ["1", "3", "4"] or EMAIL_SETTINGS["search_range"]["enabled"]:
        print("\n" + "=" * 40)
        print("Batch Processing Configuration")
        print("=" * 40)
        print("Batch size determines how many emails are processed at once.")
        print("Smaller batches use less memory but may be slower.")
        print("Larger batches are faster but use more memory.")
        
        try:
            batch_size = int(input("Enter batch size for processing (default 50): ").strip() or "50")
            EMAIL_SETTINGS["search_range"]["batch_size"] = batch_size
            print(f"✓ Batch size set to {batch_size}")
        except ValueError:
            print("Invalid batch size. Using default of 50.")
            EMAIL_SETTINGS["search_range"]["batch_size"] = 50
    else:
        print("\n✓ Configuration complete!")

def save_configuration_robust():
    """Save the configuration to email_config.py using a more robust approach"""
    print("\n=== Saving Configuration ===")
    
    config_file = "email_config.py"
    
    try:
        # Get current settings
        search_range = EMAIL_SETTINGS["search_range"]
        folder_name = EMAIL_SETTINGS["folder_name"]
        
        # Build the complete file content
        file_content = f'''import os
from dotenv import load_dotenv
# Email Configuration
# Update these settings according to your email provider

# Gmail Configuration
GMAIL_CONFIG = {{
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
}}

# Outlook/Hotmail Configuration
OUTLOOK_CONFIG = {{
    "imap_server": "outlook.office365.com",
    "imap_port": 993,
    "smtp_server": "smtp-mail.outlook.com",
    "smtp_port": 587
}}

# Yahoo Configuration
YAHOO_CONFIG = {{
    "imap_server": "imap.mail.yahoo.com",
    "imap_port": 993,
    "smtp_server": "smtp.mail.yahoo.com",
    "smtp_port": 587
}}

# QQ Mail Configuration
QQ_CONFIG = {{
    "imap_server": "imap.qq.com",
    "imap_port": 993,
    "smtp_server": "smtp.qq.com",
    "smtp_port": 587
}}

# 163 Mail Configuration
MAIL163_CONFIG = {{
    "imap_server": "imap.163.com",
    "imap_port": 993,
    "smtp_server": "smtp.163.com",
    "smtp_port": 587
}}

# Wangyi Corporate Mail Configuration
WANGYI_CONFIG = {{
    "imap_server": "imap.qiye.163.com",
    "imap_port": 993,
    "smtp_server": "smtp.qiye.163.com",
    "smtp_port": 465
}}

# Default configuration (change this to your email provider)
DEFAULT_CONFIG = WANGYI_CONFIG

# Your email settings (update these)
# Load email address, password and folder name from credentials.env
# load env
load_dotenv('credentials.env')
# read account and password
email_address = os.getenv("email_address")
password = os.getenv("password")
folder_name = os.getenv("folder_name") or "INBOX"

EMAIL_SETTINGS = {{
    "email_address": email_address,  # Replace with your email
    "password": password,  # Will prompt for password if None
    "folder_name": "{folder_name}",  # Email folder to search
    "download_path": "./files",  # Path to save attachments
    "file_types": ['.xls', '.xlsx'],  # File types to download
    "search_range": {{
        "enabled": {search_range["enabled"]},  # Set to True to use search range
        "type": "{search_range["type"]}",  # Options: "all", "date_range", "count_limit", "recent_days"
        "date_range": {{
            "start_date": "{search_range["date_range"]["start_date"]}",  # Format: YYYYMMDD
            "end_date": "{search_range["date_range"]["end_date"]}"     # Format: YYYYMMDD
        }},
        "count_limit": {search_range["count_limit"]},  # Maximum number of emails to process
        "recent_days": {search_range["recent_days"]},   # Process emails from last N days
        "batch_size": {search_range["batch_size"]}     # Number of emails to process in each batch
    }}
}}'''
        
        # Write the complete file
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write(file_content)
        
        print("✓ Configuration saved to email_config.py")
        return True
        
    except Exception as e:
        print(f"Error saving configuration: {e}")
        return False

def configure_search_settings():
    """Configure search range settings"""
    print("Email Search Range Configuration Tool")
    print("=" * 50)
    
    # Display current configuration
    display_current_config()
    
    # Ask if user wants to change configuration
    change = input("Do you want to change the search range configuration? (y/n): ").lower().strip()
    
    if change == 'y':
        print("\n" + "=" * 50)
        configure_search_range()
        
        # Show the updated configuration
        print("\n" + "=" * 50)
        print("Updated Configuration:")
        display_current_config()
        
        # Ask if user wants to save
        save = input("Do you want to save this configuration? (y/n): ").lower().strip()
        if save == 'y':
            if not save_configuration_robust():
                print("Please manually update email_config.py with the new settings.")
        else:
            print("Configuration not saved.")
    else:
        print("No changes made.")
    
    print("\nConfiguration complete!")

def configure_folder_selection():
    """Configure which folder to search"""
    print("=== Email Folder Configuration ===")
    
    # Get current folder
    current_folder = EMAIL_SETTINGS["folder_name"]
    print(f"Current folder: {current_folder}")
    
    # First check if email access is successful
    if not check_email_access():
        print("\n❌ Email access check failed. Cannot configure folder selection.")
        print("Please fix the email access issues first.")
        return False
    
    # Create a temporary downloader to get available folders
    email_address = EMAIL_SETTINGS["email_address"]
    downloader = EmailAttachmentDownloader(
        email_address=email_address,
        password=EMAIL_SETTINGS["password"],
        imap_server=DEFAULT_CONFIG["imap_server"],
        imap_port=DEFAULT_CONFIG["imap_port"]
    )
    
    # Connect to email
    print(f"\nConnecting to {email_address} to get available folders...")
    if not downloader.connect():
        print("Failed to connect to email server. Please check your credentials.")
        return False
    
    try:
        # List available folders
        print("\nAvailable folders:")
        folders = downloader.list_folders()
        
        if not folders:
            print("No folders found.")
            return False
        
        # Display folders with numbers
        for i, folder in enumerate(folders, 1):
            marker = " ← Current" if folder == current_folder else ""
            print(f"{i:2d}. {folder}{marker}")
        
        # Let user choose
        print(f"\nChoose a folder (1-{len(folders)}) or enter folder name manually:")
        choice = input("Enter choice: ").strip()
        
        selected_folder = None
        
        # Check if it's a number
        try:
            folder_index = int(choice) - 1
            if 0 <= folder_index < len(folders):
                selected_folder = folders[folder_index]
            else:
                print("Invalid folder number.")
                return False
        except ValueError:
            # User entered a folder name manually
            selected_folder = choice
            if selected_folder not in folders:
                print(f"Warning: '{selected_folder}' not found in available folders.")
                confirm = input("Use this folder anyway? (y/n): ").lower().strip()
                if confirm != 'y':
                    return False
        
        # Update the configuration
        EMAIL_SETTINGS["folder_name"] = selected_folder
        print(f"✓ Selected folder: {selected_folder}")
        
        return True
        
    except Exception as e:
        print(f"Error getting folders: {e}")
        return False
    finally:
        downloader.disconnect()

def save_folder_configuration():
    """Save the folder configuration to email_config.py"""
    print("\n=== Saving Folder Configuration ===")
    
    # Read the current email_config.py file
    config_file = "email_config.py"
    if not os.path.exists(config_file):
        print("Error: email_config.py not found")
        return False
    
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Update the folder_name in EMAIL_SETTINGS
        new_folder_name = EMAIL_SETTINGS["folder_name"]
        
        # Create the new EMAIL_SETTINGS string with updated folder_name
        search_range = EMAIL_SETTINGS["search_range"]
        new_settings = f'''EMAIL_SETTINGS = {{
    "email_address": email_address,  # Replace with your email
    "password": password,  # Will prompt for password if None
    "folder_name": "{new_folder_name}",  # Email folder to search
    "download_path": "./files",  # Path to save attachments
    "file_types": ['.xls', '.xlsx'],  # File types to download
    "search_range": {{
        "enabled": {search_range["enabled"]},  # Set to True to use search range
        "type": "{search_range["type"]}",  # Options: "all", "date_range", "count_limit", "recent_days"
        "date_range": {{
            "start_date": "{search_range["date_range"]["start_date"]}",  # Format: YYYYMMDD
            "end_date": "{search_range["date_range"]["end_date"]}"     # Format: YYYYMMDD
        }},
        "count_limit": {search_range["count_limit"]},  # Maximum number of emails to process
        "recent_days": {search_range["recent_days"]},   # Process emails from last N days
        "batch_size": {search_range["batch_size"]}     # Number of emails to process in each batch
    }}
}}'''
        
        # Find and replace the EMAIL_SETTINGS section
        # Use a more precise pattern that matches from EMAIL_SETTINGS = { to the closing }
        pattern = r'EMAIL_SETTINGS\s*=\s*\{[^}]*"search_range":\s*\{[^}]*\}[^}]*\}\s*\}'
        new_content = re.sub(pattern, new_settings, content, flags=re.DOTALL)
        
        # If the pattern didn't match, try a simpler approach
        if new_content == content:
            print("Warning: Could not find EMAIL_SETTINGS section to replace.")
            print("Please manually update the folder_name in email_config.py:")
            print(f"  folder_name: {new_folder_name}")
            return False
        
        # Write the updated content back to the file
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write(new_content)
        
        print("✓ Folder configuration saved to email_config.py")
        return True
        
    except Exception as e:
        print(f"Error saving configuration: {e}")
        return False

def configure_folder_settings():
    """Configure folder selection settings"""
    print("Email Folder Selection Tool")
    print("=" * 50)
    
    # Show current folder
    print(f"Current folder: {EMAIL_SETTINGS['folder_name']}")
    
    # Configure folder selection
    if configure_folder_selection():
        # Ask if user wants to save
        save = input("\nDo you want to save this folder configuration? (y/n): ").lower().strip()
        if save == 'y':
            if not save_configuration_robust():
                print("Please manually update email_config.py with the new folder.")
        else:
            print("Folder configuration not saved.")
    else:
        print("Folder configuration failed.")
    
    print("\nFolder configuration complete!")

def check_email_access():
    """Check if email can be successfully accessed with current credentials"""
    print("=== Email Access Check ===")
    
    # Get configuration
    email_address = EMAIL_SETTINGS["email_address"]
    password = EMAIL_SETTINGS["password"]
    
    # Check if credentials are available
    if not email_address:
        print("❌ Email address not found in configuration.")
        print("Please check your credentials.env file or email_config.py")
        return False
    
    if not password:
        print("❌ Password not found in configuration.")
        print("Please check your credentials.env file or email_config.py")
        return False
    
    print(f"Email: {email_address}")
    print(f"IMAP Server: {DEFAULT_CONFIG['imap_server']}:{DEFAULT_CONFIG['imap_port']}")
    
    # Create downloader instance
    downloader = EmailAttachmentDownloader(
        email_address=email_address,
        password=password,
        imap_server=DEFAULT_CONFIG["imap_server"],
        imap_port=DEFAULT_CONFIG["imap_port"]
    )
    
    # Test connection
    print("\nTesting email connection...")
    if not downloader.connect():
        print("❌ Failed to connect to email server.")
        print("Possible issues:")
        print("1. Incorrect email address or password")
        print("2. IMAP not enabled for your email account")
        print("3. Network connectivity issues")
        print("4. Email provider requires app password (for Gmail with 2FA)")
        return False
    
    try:
        # Test folder access
        print("✓ Successfully connected to email server")
        
        # List available folders
        print("\nFetching available folders...")
        folders = downloader.list_folders()
        
        if not folders:
            print("❌ No folders found. This might indicate an issue with folder permissions.")
            return False
        
        print(f"✓ Found {len(folders)} folders")
        
        # Check if target folder exists
        target_folder = EMAIL_SETTINGS["folder_name"]
        if target_folder in folders:
            print(f"✓ Target folder '{target_folder}' found")
        else:
            print(f"⚠️  Target folder '{target_folder}' not found in available folders")
            print("Available folders:")
            for folder in folders[:10]:  # Show first 10 folders
                print(f"  - {folder}")
            if len(folders) > 10:
                print(f"  ... and {len(folders) - 10} more")
        
        # Test folder selection
        print(f"\nTesting access to target folder '{target_folder}'...")
        status, messages = downloader.mail.select(target_folder)
        if status != 'OK':
            print(f"❌ Failed to access folder '{target_folder}': {status}")
            return False
        
        print(f"✓ Successfully accessed folder '{target_folder}'")
        
        # Get email count
        status, message_numbers = downloader.mail.search(None, 'ALL')
        if status != 'OK':
            print("❌ Failed to search emails in folder")
            return False
        
        message_list = message_numbers[0].split()
        total_emails = len(message_list)
        print(f"✓ Found {total_emails} emails in folder")
        
        if total_emails == 0:
            print("⚠️  Warning: No emails found in the target folder")
        else:
            print("✓ Email access check completed successfully!")
        
        return True
        
    except Exception as e:
        print(f"❌ Error during email access check: {e}")
        return False
    finally:
        downloader.disconnect()

def download_attachments_from_email():
    """
    Download attachments from the specified email folder
    """
    print("=== Email Attachment Downloader ===")
    
    # First check if email access is successful
    if not check_email_access():
        print("\n❌ Email access check failed. Please fix the issues above before proceeding.")
        return False
    
    # Get configuration
    email_address = EMAIL_SETTINGS["email_address"]
    folder_name = EMAIL_SETTINGS["folder_name"]
    download_path = EMAIL_SETTINGS["download_path"]
    file_types = EMAIL_SETTINGS["file_types"]
    search_range = EMAIL_SETTINGS["search_range"]
    
    # Create downloader instance
    downloader = EmailAttachmentDownloader(
        email_address=email_address,
        password=EMAIL_SETTINGS["password"],
        imap_server=DEFAULT_CONFIG["imap_server"],
        imap_port=DEFAULT_CONFIG["imap_port"]
    )
    
    # Connect to email (we already verified this works in check_email_access)
    print(f"\nConnecting to {email_address}...")
    if not downloader.connect():
        print("Failed to connect to email server. Please check your credentials.")
        return False
    
    try:
        # List available folders
        print("\nAvailable folders:")
        folders = downloader.list_folders()
        for folder in folders:
            print(f"  - {folder}")
        
        # Check if target folder exists
        if folder_name not in folders:
            print(f"\nWarning: Folder '{folder_name}' not found in available folders.")
            print("Available folders:", folders)
            use_inbox = input("Use INBOX instead? (y/n): ").lower().strip()
            if use_inbox == 'y':
                folder_name = "INBOX"
            else:
                return False
        
        # Download attachments based on search range configuration
        print(f"\nDownloading attachments from '{folder_name}'...")
        
        if search_range["enabled"]:
            print(f"Using search range: {search_range['type']}")
            downloaded_files = downloader.download_attachments_with_range(
                folder_name=folder_name,
                search_range=search_range,
                download_path=download_path,
                file_types=file_types
            )
        else:
            # Download all attachments (default behavior)
            print("Downloading all attachments (no search range specified)")
            downloaded_files = downloader.download_attachments_from_folder(
                folder_name=folder_name,
                download_path=download_path,
                file_types=file_types
            )
        
        if downloaded_files:
            print(f"\n✓ Successfully downloaded {len(downloaded_files)} files to {download_path}:")
            for file_path in downloaded_files:
                print(f"  - {os.path.basename(file_path)}")
            return True
        else:
            print("No files were downloaded.")
            return False
            
    except Exception as e:
        print(f"Error during download: {e}")
        return False
    finally:
        # Always disconnect
        downloader.disconnect()

def show_menu():
    """Show the main menu"""
    print("\nEmail Attachment Downloader")
    print("=" * 30)
    print("1. Download attachments(Download!)")
    print("2. Configure search settings")
    print("3. Configure folder settings")
    print("4. Test email access")
    print("5. Show current configuration")
    print("6. Exit")
    print("=" * 30)

def main():
    """
    Main function with menu system
    """
    print("Starting Email Attachment Downloader...")
    
    while True:
        show_menu()
        choice = input("Enter your choice (1-6): ").strip()
        
        if choice == "1":
            # Download attachments
            download_success = download_attachments_from_email()
            
            if download_success:
                print("\n✓ Download completed successfully!")
                print(f"Files saved to: {EMAIL_SETTINGS['download_path']}")
            else:
                print("\n✗ Download failed.")
                
        elif choice == "2":
            # Configure search settings
            configure_search_settings()
            
        elif choice == "3":
            # Configure folder settings
            configure_folder_settings()
            
        elif choice == "4":
            # Test email access
            check_email_access()
            
        elif choice == "5":
            # Show current configuration
            display_current_config()
            
        elif choice == "6":
            # Exit
            print("Goodbye!")
            break
            
        else:
            print("Invalid choice. Please enter 1-6.")
        
        if choice != "6":
            input("\nPress Enter to continue...")

if __name__ == "__main__":
    main() 