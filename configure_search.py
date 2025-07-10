#!/usr/bin/env python3
"""
Email Search Range Configuration Tool
This script helps you configure the email search range settings
"""

import os
import re
from email_config import EMAIL_SETTINGS

def display_current_config():
    """Display current search range configuration"""
    print("=== Current Email Search Range Configuration ===")
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
        start_date = input("Start date (e.g., 20240101): ").strip()
        end_date = input("End date (e.g., 20241231): ").strip()
        
        EMAIL_SETTINGS["search_range"]["date_range"]["start_date"] = start_date
        EMAIL_SETTINGS["search_range"]["date_range"]["end_date"] = end_date
        
        print(f"✓ Configured date range: {start_date} to {end_date}")
        
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
    
    # Configure batch size
    try:
        batch_size = int(input("\nEnter batch size for processing (default 50): ").strip() or "50")
        EMAIL_SETTINGS["search_range"]["batch_size"] = batch_size
        print(f"✓ Batch size set to {batch_size}")
    except ValueError:
        print("Invalid batch size. Using default of 50.")
        EMAIL_SETTINGS["search_range"]["batch_size"] = 50

def save_configuration():
    """Save the configuration to email_config.py"""
    print("\n=== Saving Configuration ===")
    
    # Read the current email_config.py file
    config_file = "email_config.py"
    if not os.path.exists(config_file):
        print("Error: email_config.py not found")
        return False
    
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Update the EMAIL_SETTINGS section
        search_range = EMAIL_SETTINGS["search_range"]
        
        # Create the new EMAIL_SETTINGS string
        new_settings = f'''EMAIL_SETTINGS = {{
    "email_address": email_address,  # Replace with your email
    "password": password,  # Will prompt for password if None
    "folder_name": folder_name,  # Email folder to search
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
        # Use a more specific pattern to match the entire EMAIL_SETTINGS block
        pattern = r'EMAIL_SETTINGS\s*=\s*\{[^}]*"search_range":\s*\{[^}]*\}[^}]*\}'
        new_content = re.sub(pattern, new_settings, content, flags=re.DOTALL)
        
        # If the pattern didn't match, try a simpler approach
        if new_content == content:
            print("Warning: Could not find EMAIL_SETTINGS section to replace.")
            print("Please manually update the search_range settings in email_config.py:")
            print(f"  enabled: {search_range['enabled']}")
            print(f"  type: {search_range['type']}")
            print(f"  batch_size: {search_range['batch_size']}")
            return False
        
        # Write the updated content back to the file
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write(new_content)
        
        print("✓ Configuration saved to email_config.py")
        return True
        
    except Exception as e:
        print(f"Error saving configuration: {e}")
        return False

def main():
    """Main configuration function"""
    print("Email Search Range Configuration Tool")
    print("=" * 50)
    
    # Display current configuration
    display_current_config()
    
    # Ask if user wants to change configuration
    change = input("Do you want to change the search range configuration? (y/n): ").lower().strip()
    
    if change == 'y':
        configure_search_range()
        
        # Ask if user wants to save
        save = input("\nDo you want to save this configuration? (y/n): ").lower().strip()
        if save == 'y':
            if not save_configuration():
                print("Please manually update email_config.py with the new settings.")
        else:
            print("Configuration not saved.")
    else:
        print("No changes made.")
    
    print("\nConfiguration complete!")

if __name__ == "__main__":
    main() 