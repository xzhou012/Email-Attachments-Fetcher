import imaplib
import email
import os
import getpass
from email.header import decode_header
from datetime import datetime
import pandas as pd

class EmailAttachmentDownloader:
    def __init__(self, email_address, password=None, imap_server="imap.gmail.com", imap_port=993):
        """
        Initialize the email downloader
        
        Args:
            email_address (str): Your email address
            password (str): Your email password or app password
            imap_server (str): IMAP server address
            imap_port (int): IMAP server port
        """
        self.email_address = email_address
        self.password = password or getpass.getpass("Enter your email password: ")
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.mail = None
        
    def connect(self):
        """Connect to the email server"""
        try:
            self.mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            self.mail.login(self.email_address, self.password)
            print(f"Successfully connected to {self.email_address}")
            return True
        except Exception as e:
            print(f"Failed to connect: {e}")
            return False
    
    def disconnect(self):
        """Disconnect from the email server"""
        if self.mail:
            self.mail.logout()
            print("Disconnected from email server")
    
    def _safe_decode(self, text, default_encoding='utf-8'):
        """
        Safely decode text with fallback encodings
        """
        if isinstance(text, str):
            return text
        
        if not isinstance(text, bytes):
            return str(text)
        
        # Try different encodings
        encodings = [default_encoding, 'gbk', 'gb2312', 'big5', 'latin1', 'cp1252']
        
        for encoding in encodings:
            try:
                return text.decode(encoding)
            except (UnicodeDecodeError, LookupError):
                continue
        
        # If all encodings fail, use 'replace' to handle invalid characters
        try:
            return text.decode(default_encoding, errors='replace')
        except:
            return str(text)
    
    def _decode_filename(self, filename):
        """
        Decode MIME-encoded filename properly
        """
        if not filename:
            return ""
        
        try:
            # Use email.header.decode_header to properly decode MIME-encoded filenames
            from email.header import decode_header
            decoded_parts = decode_header(filename)
            
            # Combine all decoded parts
            decoded_filename = ""
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        try:
                            decoded_filename += part.decode(encoding)
                        except:
                            decoded_filename += self._safe_decode(part)
                    else:
                        decoded_filename += self._safe_decode(part)
                else:
                    decoded_filename += str(part)
            
            return decoded_filename.strip()
            
        except Exception as e:
            # Fallback to simple decoding
            return self._safe_decode(filename)
    
    def list_folders(self):
        """List all available folders"""
        if not self.mail:
            print("Not connected to email server")
            return []
        
        try:
            status, folders = self.mail.list()
            folder_list = []
            for folder in folders:
                try:
                    folder_name = folder.decode()
                    # Handle different folder name formats
                    if '"' in folder_name:
                        folder_name = folder_name.split('"')[-2]
                    else:
                        folder_name = folder_name.split()[-1]
                    
                    # Decode the folder name safely
                    folder_name = self._safe_decode(folder_name.encode('latin1') if isinstance(folder_name, str) else folder_name)
                    folder_list.append(folder_name)
                except Exception as e:
                    print(f"Error decoding folder name: {e}")
                    continue
            return folder_list
        except Exception as e:
            print(f"Error listing folders: {e}")
            return []
    
    def download_attachments_from_folder(self, folder_name, download_path="./files", file_types=None, batch_size=100):
        """
        Download all attachments from a specific folder
        
        Args:
            folder_name (str): Name of the email folder
            download_path (str): Path to save attachments
            file_types (list): List of file extensions to download (e.g., ['.xls', '.xlsx', '.pdf'])
            batch_size (int): Number of emails to process in each batch
        """
        if not self.mail:
            print("Not connected to email server")
            return
        
        # Create download directory if it doesn't exist
        if not os.path.exists(download_path):
            os.makedirs(download_path)
            print(f"Created directory: {download_path}")
        
        try:
            # Select the folder
            status, messages = self.mail.select(folder_name)
            if status != 'OK':
                print(f"Failed to select folder: {folder_name}")
                return
            
            # Get total number of emails in the folder
            status, message_numbers = self.mail.search(None, 'ALL')
            if status != 'OK':
                print("Failed to search emails")
                return
            
            message_list = message_numbers[0].split()
            total_emails = len(message_list)
            downloaded_files = []
            
            print(f"Found {total_emails} emails in folder '{folder_name}'")
            print(f"Processing in batches of {batch_size} emails...")
            
            # Process emails in reverse order (oldest first) to ensure we get all emails
            # Convert to list and reverse to process from oldest to newest
            message_list = list(message_list)
            message_list.reverse()
            
            # Process emails in batches
            for batch_start in range(0, len(message_list), batch_size):
                batch_end = min(batch_start + batch_size, len(message_list))
                batch_messages = message_list[batch_start:batch_end]
                
                print(f"\n--- Processing batch {batch_start//batch_size + 1}/{(len(message_list)-1)//batch_size + 1} (emails {batch_start+1}-{batch_end}) ---")
                
                for i, message_num in enumerate(batch_messages, batch_start + 1):
                    print(f"Processing email {i}/{total_emails} (Message ID: {message_num.decode()})")
                    
                    try:
                        # Fetch the email
                        status, msg_data = self.mail.fetch(message_num, '(RFC822)')
                        if status != 'OK':
                            print(f"  ✗ Failed to fetch email {message_num.decode()}")
                            continue
                        
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)
                        
                        # Get email subject safely
                        subject = email_message["subject"]
                        if subject:
                            try:
                                decoded_subject = decode_header(subject)
                                subject = ''.join([self._safe_decode(text) if isinstance(text, bytes) else str(text) 
                                                 for text, encoding in decoded_subject])
                            except:
                                subject = self._safe_decode(subject)
                        else:
                            subject = "No Subject"
                        
                        print(f"  Subject: {subject[:50]}...")
                        
                        # Process attachments
                        email_attachments = 0
                        for part in email_message.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            if part.get('Content-Disposition') is None:
                                continue
                            
                            filename = part.get_filename()
                            if filename:
                                # Decode filename safely
                                filename = self._decode_filename(filename)
                                
                                # Check file type filter
                                if file_types:
                                    file_ext = os.path.splitext(filename)[1].lower()
                                    if file_ext not in file_types:
                                        print(f"    Skipping {filename} (not in allowed types)")
                                        continue
                                
                                # Create safe filename
                                safe_filename = self._create_safe_filename(filename)
                                file_path = os.path.join(download_path, safe_filename)
                                
                                # Check if file already exists
                                if os.path.exists(file_path):
                                    print(f"    File already exists: {safe_filename}")
                                    downloaded_files.append(file_path)
                                    email_attachments += 1
                                    continue
                                
                                # Save attachment
                                try:
                                    with open(file_path, 'wb') as f:
                                        f.write(part.get_payload(decode=True))
                                    
                                    downloaded_files.append(file_path)
                                    email_attachments += 1
                                    print(f"    ✓ Downloaded: {safe_filename}")
                                except Exception as e:
                                    print(f"    ✗ Error saving {safe_filename}: {e}")
                                    continue
                        
                        if email_attachments > 0:
                            print(f"  ✓ Downloaded {email_attachments} attachment(s) from this email")
                        else:
                            print(f"  - No matching attachments in this email")
                    
                    except Exception as e:
                        print(f"  ✗ Error processing email {message_num.decode()}: {e}")
                        continue
                
                print(f"--- Completed batch {batch_start//batch_size + 1} ---")
            
            print(f"\nDownloaded {len(downloaded_files)} files to {download_path}")
            return downloaded_files
            
        except Exception as e:
            print(f"Error downloading attachments: {e}")
            return []
    
    def _create_safe_filename(self, filename):
        """Create a safe filename by removing invalid characters"""
        if not filename:
            return "unnamed_file"
        
        # Remove or replace invalid characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        # Remove leading/trailing spaces and dots
        filename = filename.strip(' .')
        
        # Ensure filename is not too long
        if len(filename) > 200:
            name, ext = os.path.splitext(filename)
            filename = name[:200-len(ext)] + ext
        
        # Ensure filename is not empty
        if not filename:
            filename = "unnamed_file"
        
        return filename
    
    def _convert_date_to_imap_format(self, date_str):
        """
        Convert YYYYMMDD format to IMAP-compatible DD-MMM-YYYY format
        
        Args:
            date_str (str): Date in YYYYMMDD format
            
        Returns:
            str: Date in DD-MMM-YYYY format for IMAP
        """
        try:
            from datetime import datetime
            # Parse YYYYMMDD format
            date_obj = datetime.strptime(date_str, '%Y%m%d')
            # Convert to DD-MMM-YYYY format
            return date_obj.strftime('%d-%b-%Y')
        except ValueError as e:
            print(f"Error converting date {date_str}: {e}")
            return date_str  # Return original if conversion fails
    
    def download_attachments_by_date_range(self, folder_name, start_date, end_date, download_path="./files", file_types=None):
        """
        Download attachments from emails within a specific date range
        
        Args:
            folder_name (str): Name of the email folder
            start_date (str): Start date in format 'YYYYMMDD'
            end_date (str): End date in format 'YYYYMMDD'
            download_path (str): Path to save attachments
            file_types (list): List of file extensions to download
        """
        if not self.mail:
            print("Not connected to email server")
            return
        
        # Create download directory if it doesn't exist
        if not os.path.exists(download_path):
            os.makedirs(download_path)
        
        try:
            # Select the folder
            status, messages = self.mail.select(folder_name)
            if status != 'OK':
                print(f"Failed to select folder: {folder_name}")
                return
            
            # Convert dates to IMAP format
            imap_start_date = self._convert_date_to_imap_format(start_date)
            imap_end_date = self._convert_date_to_imap_format(end_date)
            
            print(f"Searching for emails from {imap_start_date} to {imap_end_date}")
            
            # Search for emails within date range
            search_criteria = f'(SINCE "{imap_start_date}" BEFORE "{imap_end_date}")'
            status, message_numbers = self.mail.search(None, search_criteria)
            if status != 'OK':
                print("Failed to search emails")
                return
            
            message_list = message_numbers[0].split()
            total_emails = len(message_list)
            downloaded_files = []
            
            print(f"Found {total_emails} emails in folder '{folder_name}' between {start_date} and {end_date}")
            
            for i, message_num in enumerate(message_list, 1):
                print(f"Processing email {i}/{total_emails}")
                
                try:
                    # Fetch the email
                    status, msg_data = self.mail.fetch(message_num, '(RFC822)')
                    if status != 'OK':
                        continue
                    
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)
                    
                    # Get email subject safely
                    subject = email_message["subject"]
                    if subject:
                        try:
                            decoded_subject = decode_header(subject)
                            subject = ''.join([self._safe_decode(text) if isinstance(text, bytes) else str(text) 
                                             for text, encoding in decoded_subject])
                        except:
                            subject = self._safe_decode(subject)
                    else:
                        subject = "No Subject"
                    
                    print(f"  Subject: {subject}")
                    
                    # Process attachments
                    for part in email_message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        
                        filename = part.get_filename()
                        if filename:
                            # Decode filename safely
                            filename = self._decode_filename(filename)
                            
                            # Check file type filter
                            if file_types:
                                file_ext = os.path.splitext(filename)[1].lower()
                                if file_ext not in file_types:
                                    continue
                            
                            # Create safe filename
                            safe_filename = self._create_safe_filename(filename)
                            file_path = os.path.join(download_path, safe_filename)
                            
                            # Save attachment
                            try:
                                with open(file_path, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                
                                downloaded_files.append(file_path)
                                print(f"    Downloaded: {safe_filename}")
                            except Exception as e:
                                print(f"    Error saving {safe_filename}: {e}")
                                continue
                
                except Exception as e:
                    print(f"  Error processing email {i}: {e}")
                    continue
            
            print(f"\nDownloaded {len(downloaded_files)} files to {download_path}")
            return downloaded_files
            
        except Exception as e:
            print(f"Error downloading attachments: {e}")
            return []

    def count_emails_with_attachments(self, folder_name, file_types=None):
        """
        Count emails with attachments in a folder
        
        Args:
            folder_name (str): Name of the email folder
            file_types (list): List of file extensions to count
        """
        if not self.mail:
            print("Not connected to email server")
            return 0
        
        try:
            # Select the folder
            status, messages = self.mail.select(folder_name)
            if status != 'OK':
                print(f"Failed to select folder: {folder_name}")
                return 0
            
            # Get total number of emails in the folder
            status, message_numbers = self.mail.search(None, 'ALL')
            if status != 'OK':
                print("Failed to search emails")
                return 0
            
            message_list = message_numbers[0].split()
            total_emails = len(message_list)
            emails_with_attachments = 0
            total_attachments = 0
            
            print(f"Scanning {total_emails} emails for attachments...")
            
            # Process emails in reverse order (oldest first)
            message_list = list(message_list)
            message_list.reverse()
            
            for i, message_num in enumerate(message_list, 1):
                if i % 50 == 0:  # Progress update every 50 emails
                    print(f"Scanned {i}/{total_emails} emails...")
                
                try:
                    # Fetch the email
                    status, msg_data = self.mail.fetch(message_num, '(RFC822)')
                    if status != 'OK':
                        continue
                    
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)
                    
                    # Check for attachments
                    email_has_attachments = False
                    for part in email_message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        
                        filename = part.get_filename()
                        if filename:
                            # Decode filename safely
                            filename = self._decode_filename(filename)
                            
                            # Check file type filter
                            if file_types:
                                file_ext = os.path.splitext(filename)[1].lower()
                                if file_ext in file_types:
                                    email_has_attachments = True
                                    total_attachments += 1
                            else:
                                email_has_attachments = True
                                total_attachments += 1
                    
                    if email_has_attachments:
                        emails_with_attachments += 1
                
                except Exception as e:
                    continue
            
            print(f"Found {emails_with_attachments} emails with attachments out of {total_emails} total emails")
            print(f"Total attachments found: {total_attachments}")
            return emails_with_attachments
            
        except Exception as e:
            print(f"Error counting emails: {e}")
            return 0

    def download_attachments_with_range(self, folder_name, search_range, download_path="./files", file_types=None):
        """
        Download attachments with various search range options
        
        Args:
            folder_name (str): Name of the email folder
            search_range (dict): Search range configuration
            download_path (str): Path to save attachments
            file_types (list): List of file extensions to download
        """
        if not self.mail:
            print("Not connected to email server")
            return
        
        search_type = search_range.get("type", "all")
        
        if search_type == "all":
            return self.download_attachments_from_folder(folder_name, download_path, file_types, search_range.get("batch_size", 100))
        elif search_type == "date_range":
            date_range = search_range.get("date_range", {})
            return self.download_attachments_by_date_range(
                folder_name, 
                date_range.get("start_date"), 
                date_range.get("end_date"), 
                download_path, 
                file_types
            )
        elif search_type == "count_limit":
            return self.download_attachments_with_count_limit(
                folder_name, 
                search_range.get("count_limit", 100), 
                download_path, 
                file_types, 
                search_range.get("batch_size", 50)
            )
        elif search_type == "recent_days":
            return self.download_attachments_recent_days(
                folder_name, 
                search_range.get("recent_days", 30), 
                download_path, 
                file_types, 
                search_range.get("batch_size", 50)
            )
        else:
            print(f"Unknown search type: {search_type}")
            return []
    
    def download_attachments_with_count_limit(self, folder_name, count_limit, download_path="./files", file_types=None, batch_size=50):
        """
        Download attachments from a limited number of emails
        
        Args:
            folder_name (str): Name of the email folder
            count_limit (int): Maximum number of emails to process
            download_path (str): Path to save attachments
            file_types (list): List of file extensions to download
            batch_size (int): Number of emails to process in each batch
        """
        if not self.mail:
            print("Not connected to email server")
            return
        
        # Create download directory if it doesn't exist
        if not os.path.exists(download_path):
            os.makedirs(download_path)
            print(f"Created directory: {download_path}")
        
        try:
            # Select the folder
            status, messages = self.mail.select(folder_name)
            if status != 'OK':
                print(f"Failed to select folder: {folder_name}")
                return
            
            # Get total number of emails in the folder
            status, message_numbers = self.mail.search(None, 'ALL')
            if status != 'OK':
                print("Failed to search emails")
                return
            
            message_list = message_numbers[0].split()
            total_emails = len(message_list)
            downloaded_files = []
            
            print(f"Found {total_emails} emails in folder '{folder_name}'")
            print(f"Processing first {count_limit} emails in batches of {batch_size}...")
            
            # Process emails in reverse order (oldest first) to ensure we get all emails
            # Convert to list and reverse to process from oldest to newest
            message_list = list(message_list)
            message_list.reverse()
            
            # Limit to the specified count
            message_list = message_list[:count_limit]
            actual_count = len(message_list)
            
            # Process emails in batches
            for batch_start in range(0, len(message_list), batch_size):
                batch_end = min(batch_start + batch_size, len(message_list))
                batch_messages = message_list[batch_start:batch_end]
                
                print(f"\n--- Processing batch {batch_start//batch_size + 1}/{(len(message_list)-1)//batch_size + 1} (emails {batch_start+1}-{batch_end} of {actual_count}) ---")
                
                for i, message_num in enumerate(batch_messages, batch_start + 1):
                    print(f"Processing email {i}/{actual_count} (Message ID: {message_num.decode()})")
                    
                    try:
                        # Fetch the email
                        status, msg_data = self.mail.fetch(message_num, '(RFC822)')
                        if status != 'OK':
                            print(f"  ✗ Failed to fetch email {message_num.decode()}")
                            continue
                        
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)
                        
                        # Get email subject safely
                        subject = email_message["subject"]
                        if subject:
                            try:
                                decoded_subject = decode_header(subject)
                                subject = ''.join([self._safe_decode(text) if isinstance(text, bytes) else str(text) 
                                                 for text, encoding in decoded_subject])
                            except:
                                subject = self._safe_decode(subject)
                        else:
                            subject = "No Subject"
                        
                        print(f"  Subject: {subject[:50]}...")
                        
                        # Process attachments
                        email_attachments = 0
                        for part in email_message.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            if part.get('Content-Disposition') is None:
                                continue
                            
                            filename = part.get_filename()
                            if filename:
                                # Decode filename safely
                                filename = self._decode_filename(filename)
                                
                                # Check file type filter
                                if file_types:
                                    file_ext = os.path.splitext(filename)[1].lower()
                                    if file_ext not in file_types:
                                        print(f"    Skipping {filename} (not in allowed types)")
                                        continue
                                
                                # Create safe filename
                                safe_filename = self._create_safe_filename(filename)
                                file_path = os.path.join(download_path, safe_filename)
                                
                                # Check if file already exists
                                if os.path.exists(file_path):
                                    print(f"    File already exists: {safe_filename}")
                                    downloaded_files.append(file_path)
                                    email_attachments += 1
                                    continue
                                
                                # Save attachment
                                try:
                                    with open(file_path, 'wb') as f:
                                        f.write(part.get_payload(decode=True))
                                    
                                    downloaded_files.append(file_path)
                                    email_attachments += 1
                                    print(f"    ✓ Downloaded: {safe_filename}")
                                except Exception as e:
                                    print(f"    ✗ Error saving {safe_filename}: {e}")
                                    continue
                        
                        if email_attachments > 0:
                            print(f"  ✓ Downloaded {email_attachments} attachment(s) from this email")
                        else:
                            print(f"  - No matching attachments in this email")
                    
                    except Exception as e:
                        print(f"  ✗ Error processing email {message_num.decode()}: {e}")
                        continue
                
                print(f"--- Completed batch {batch_start//batch_size + 1} ---")
            
            print(f"\nDownloaded {len(downloaded_files)} files to {download_path}")
            return downloaded_files
            
        except Exception as e:
            print(f"Error downloading attachments: {e}")
            return []
    
    def download_attachments_recent_days(self, folder_name, days, download_path="./files", file_types=None, batch_size=50):
        """
        Download attachments from emails received in the last N days
        
        Args:
            folder_name (str): Name of the email folder
            days (int): Number of recent days to search
            download_path (str): Path to save attachments
            file_types (list): List of file extensions to download
            batch_size (int): Number of emails to process in each batch
        """
        if not self.mail:
            print("Not connected to email server")
            return
        
        # Create download directory if it doesn't exist
        if not os.path.exists(download_path):
            os.makedirs(download_path)
            print(f"Created directory: {download_path}")
        
        try:
            # Select the folder
            status, messages = self.mail.select(folder_name)
            if status != 'OK':
                print(f"Failed to select folder: {folder_name}")
                return
            
            # Calculate the date for N days ago
            from datetime import datetime, timedelta
            end_date = datetime.now()
            start_date = end_date - timedelta(days=days)
            
            # Format dates for IMAP search
            start_date_str = start_date.strftime("%d-%b-%Y")
            end_date_str = end_date.strftime("%d-%b-%Y")
            
            print(f"Searching for emails from {start_date_str} to {end_date_str} (last {days} days)")
            
            # Search for emails within date range
            search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
            status, message_numbers = self.mail.search(None, search_criteria)
            if status != 'OK':
                print("Failed to search emails")
                return
            
            message_list = message_numbers[0].split()
            total_emails = len(message_list)
            downloaded_files = []
            
            print(f"Found {total_emails} emails in the last {days} days")
            
            if total_emails == 0:
                print("No emails found in the specified date range")
                return []
            
            # Process emails in reverse order (oldest first)
            message_list = list(message_list)
            message_list.reverse()
            
            # Process emails in batches
            for batch_start in range(0, len(message_list), batch_size):
                batch_end = min(batch_start + batch_size, len(message_list))
                batch_messages = message_list[batch_start:batch_end]
                
                print(f"\n--- Processing batch {batch_start//batch_size + 1}/{(len(message_list)-1)//batch_size + 1} (emails {batch_start+1}-{batch_end}) ---")
                
                for i, message_num in enumerate(batch_messages, batch_start + 1):
                    print(f"Processing email {i}/{total_emails} (Message ID: {message_num.decode()})")
                    
                    try:
                        # Fetch the email
                        status, msg_data = self.mail.fetch(message_num, '(RFC822)')
                        if status != 'OK':
                            print(f"  ✗ Failed to fetch email {message_num.decode()}")
                            continue
                        
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)
                        
                        # Get email subject safely
                        subject = email_message["subject"]
                        if subject:
                            try:
                                decoded_subject = decode_header(subject)
                                subject = ''.join([self._safe_decode(text) if isinstance(text, bytes) else str(text) 
                                                 for text, encoding in decoded_subject])
                            except:
                                subject = self._safe_decode(subject)
                        else:
                            subject = "No Subject"
                        
                        print(f"  Subject: {subject[:50]}...")
                        
                        # Process attachments
                        email_attachments = 0
                        for part in email_message.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            if part.get('Content-Disposition') is None:
                                continue
                            
                            filename = part.get_filename()
                            if filename:
                                # Decode filename safely
                                filename = self._decode_filename(filename)
                                
                                # Check file type filter
                                if file_types:
                                    file_ext = os.path.splitext(filename)[1].lower()
                                    if file_ext not in file_types:
                                        print(f"    Skipping {filename} (not in allowed types)")
                                        continue
                                
                                # Create safe filename
                                safe_filename = self._create_safe_filename(filename)
                                file_path = os.path.join(download_path, safe_filename)
                                
                                # Check if file already exists
                                if os.path.exists(file_path):
                                    print(f"    File already exists: {safe_filename}")
                                    downloaded_files.append(file_path)
                                    email_attachments += 1
                                    continue
                                
                                # Save attachment
                                try:
                                    with open(file_path, 'wb') as f:
                                        f.write(part.get_payload(decode=True))
                                    
                                    downloaded_files.append(file_path)
                                    email_attachments += 1
                                    print(f"    ✓ Downloaded: {safe_filename}")
                                except Exception as e:
                                    print(f"    ✗ Error saving {safe_filename}: {e}")
                                    continue
                        
                        if email_attachments > 0:
                            print(f"  ✓ Downloaded {email_attachments} attachment(s) from this email")
                        else:
                            print(f"  - No matching attachments in this email")
                    
                    except Exception as e:
                        print(f"  ✗ Error processing email {message_num.decode()}: {e}")
                        continue
                
                print(f"--- Completed batch {batch_start//batch_size + 1} ---")
            
            print(f"\nDownloaded {len(downloaded_files)} files to {download_path}")
            return downloaded_files
            
        except Exception as e:
            print(f"Error downloading attachments: {e}")
            return []

def main():
    """Example usage of the EmailAttachmentDownloader"""
    
    # Configuration
    EMAIL_ADDRESS = "your_email@gmail.com"  # Replace with your email
    FOLDER_NAME = "INBOX"  # Replace with your target folder name
    DOWNLOAD_PATH = "./files"  # Path to save attachments
    FILE_TYPES = ['.xls', '.xlsx', '.pdf', '.doc', '.docx']  # File types to download
    
    # Create downloader instance
    downloader = EmailAttachmentDownloader(EMAIL_ADDRESS)
    
    # Connect to email
    if not downloader.connect():
        return
    
    try:
        # List available folders
        print("\nAvailable folders:")
        folders = downloader.list_folders()
        for folder in folders:
            print(f"  - {folder}")
        
        # Download all attachments from the specified folder
        print(f"\nDownloading attachments from '{FOLDER_NAME}'...")
        downloaded_files = downloader.download_attachments_from_folder(
            folder_name=FOLDER_NAME,
            download_path=DOWNLOAD_PATH,
            file_types=FILE_TYPES
        )
        
        if downloaded_files:
            print(f"\nSuccessfully downloaded {len(downloaded_files)} files:")
            for file_path in downloaded_files:
                print(f"  - {os.path.basename(file_path)}")
        else:
            print("No files were downloaded.")
            
    finally:
        # Always disconnect
        downloader.disconnect()

if __name__ == "__main__":
    main() 