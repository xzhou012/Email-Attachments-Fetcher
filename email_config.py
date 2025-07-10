import os
from dotenv import load_dotenv
# Email Configuration
# Update these settings according to your email provider

# Gmail Configuration
GMAIL_CONFIG = {
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
}

# Outlook/Hotmail Configuration
OUTLOOK_CONFIG = {
    "imap_server": "outlook.office365.com",
    "imap_port": 993,
    "smtp_server": "smtp-mail.outlook.com",
    "smtp_port": 587
}

# Yahoo Configuration
YAHOO_CONFIG = {
    "imap_server": "imap.mail.yahoo.com",
    "imap_port": 993,
    "smtp_server": "smtp.mail.yahoo.com",
    "smtp_port": 587
}

# QQ Mail Configuration
QQ_CONFIG = {
    "imap_server": "imap.qq.com",
    "imap_port": 993,
    "smtp_server": "smtp.qq.com",
    "smtp_port": 587
}

# 163 Mail Configuration
MAIL163_CONFIG = {
    "imap_server": "imap.163.com",
    "imap_port": 993,
    "smtp_server": "smtp.163.com",
    "smtp_port": 587
}

# Wangyi Corporate Mail Configuration
WANGYI_CONFIG = {
    "imap_server": "imap.qiye.163.com",
    "imap_port": 993,
    "smtp_server": "smtp.qiye.163.com",
    "smtp_port": 465
}

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

EMAIL_SETTINGS = {
    "email_address": email_address,  # Replace with your email
    "password": password,  # Will prompt for password if None
    "folder_name": "INBOX/&TzBQPIho-",  # Email folder to search
    "download_path": "./files",  # Path to save attachments
    "file_types": ['.xls', '.xlsx'],  # File types to download
    "search_range": {
        "enabled": True,  # Set to True to use search range
        "type": "recent_days",  # Options: "all", "date_range", "count_limit", "recent_days"
        "date_range": {
            "start_date": "20250615",  # Format: YYYYMMDD
            "end_date": "20250630"     # Format: YYYYMMDD
        },
        "count_limit": 100,  # Maximum number of emails to process
        "recent_days": 7,   # Process emails from last N days
        "batch_size": 50     # Number of emails to process in each batch
    }
}