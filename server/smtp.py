# server/smtp.py
# SMTP Configuration and Email Service for MERQ Timesheet System
# Version: 1.0.0
# Author: Michael Kifle Teferra
# Date: November 2025

import os
import sys
import smtplib
import logging
import tempfile
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import unicodedata

# Add the src directory to Python path to import necessary modules
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

try:
    from timesheet import EthiopianDateConverter
except ImportError:
    # Fallback: Define a minimal EthiopianDateConverter if import fails
    class EthiopianDateConverter:
        @staticmethod
        def gregorian_to_ethiopian(greg_date):
            # Simplified implementation
            return 2017, 2, 27
        
        @staticmethod 
        def get_current_ethiopian_date():
            return {
                'year': 2017,
                'month': 2,
                'day': 27,
                'month_name': 'ጥቅምት'
            }

class SMTPConfig:
    """SMTP Configuration Manager"""
    
    def __init__(self):
        self.config = {
            'SMTPServer': 'cloud.merqconsultancy.org',
            'SMTPPort': 587,
            'SMTPUser': 'app@cloud.merqconsultancy.org',
            'SMTPPassword': 'MerqAppCloud',
            'UseTLS': False
        }
        
        # Setup logging
        self.setup_logging()
    
    def setup_logging(self):
        """Setup logging configuration"""
        log_dir = 'server'
        os.makedirs(log_dir, exist_ok=True)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(os.path.join(log_dir, 'smtp.log'), encoding='utf-8'),
                logging.StreamHandler(sys.stdout)  # Use stdout which supports Unicode
            ]
        )
        
        self.logger = logging.getLogger('MERQ_SMTP')
    
    def get_config(self):
        """Get SMTP configuration"""
        return self.config
    
    def update_config(self, new_config):
        """Update SMTP configuration"""
        self.config.update(new_config)
        self.logger.info("SMTP configuration updated")

class EmailService:
    """Email Service for sending timesheet emails"""
    
    def __init__(self):
        self.smtp_config = SMTPConfig()
        self.config = self.smtp_config.get_config()
        self.logger = self.smtp_config.logger
    
    def send_timesheet_email(self, timesheet_file_path, user_session, hr_users, selected_month=None, selected_year=None):
        """
        Send timesheet via email to HR with CC to user
        
        Args:
            timesheet_file_path: Path to the generated Excel file
            user_session: UserSession object with user data
            hr_users: List of HR users
            selected_month: The month the timesheet was generated for
            selected_year: The year the timesheet was generated for
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            self.logger.info(f"Starting email send process")
            self.logger.info(f"File path: {timesheet_file_path}")
            self.logger.info(f"User: {user_session.full_name if user_session else 'None'}")
            self.logger.info(f"Selected month: {selected_month}, year: {selected_year}")
            
            # Verify the timesheet file exists
            if not timesheet_file_path:
                self.logger.error("No timesheet file path provided")
                return False
                
            if not os.path.exists(timesheet_file_path):
                self.logger.error(f"Timesheet file not found: {timesheet_file_path}")
                self.logger.error(f"Current directory: {os.getcwd()}")
                return False
            
            # Verify file is not empty
            file_size = os.path.getsize(timesheet_file_path)
            if file_size == 0:
                self.logger.error(f"Timesheet file is empty: {timesheet_file_path}")
                return False
            
            self.logger.info(f"File exists and has size: {file_size} bytes")
            
            # Use selected month/year if provided, otherwise use current
            if selected_month and selected_year:
                month_name = selected_month
                year = selected_year
                self.logger.info(f"Using selected period: {month_name} {year}")
            else:
                # Fallback to current Ethiopian date
                eth_converter = EthiopianDateConverter()
                current_eth_date = eth_converter.get_current_ethiopian_date()
                month_name = current_eth_date['month_name']
                year = current_eth_date['year']
                self.logger.info(f"Using current period: {month_name} {year}")
            
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.config['SMTPUser']
            
            # Set recipients - HR users and CC the sender
            hr_emails = [hr['email'] for hr in hr_users if hr.get('email')]
            if not hr_emails:
                hr_emails = ['support@merqconsultancy.org']  # Fallback
            
            msg['To'] = ', '.join(hr_emails)
            msg['Cc'] = user_session.email if user_session else 'unknown@merqconsultancy.org'
            msg['Subject'] = f"MERQ Timesheet for {month_name} {year} - {user_session.full_name if user_session else 'Unknown User'}"
            
            self.logger.info(f"Email recipients - To: {hr_emails}, Cc: {msg['Cc']}")
            
            # Email body with better formatting
            body = self._create_email_body(user_session, month_name, year, hr_users, timesheet_file_path)
            msg.attach(MIMEText(body, 'html'))
            
            # Attach timesheet file with ASCII-only filename to avoid "noname" issue
            try:
                with open(timesheet_file_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                
                # Create ASCII-only filename to avoid "noname" issue in email clients
                original_filename = os.path.basename(timesheet_file_path)
                safe_filename = self._create_safe_filename(original_filename, user_session, month_name, year)
                
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{safe_filename}"'
                )
                part.add_header(
                    'Content-Type',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                msg.attach(part)
                self.logger.info(f"Successfully attached Excel file: {safe_filename} (original: {original_filename})")
                
            except Exception as e:
                self.logger.error(f"Error attaching file: {e}")
                return False
            
            # Send email
            all_recipients = hr_emails + [msg['Cc']]
            success = self._send_email(msg, all_recipients)
            
            if success:
                self.logger.info(f"Timesheet email sent successfully for {user_session.full_name if user_session else 'Unknown User'}")
                self.logger.info(f"Recipients: HR={hr_emails}, CC={msg['Cc']}")
                self.logger.info(f"Attachment: {safe_filename}")
                return True
            else:
                self.logger.error(f"Failed to send timesheet email for {user_session.full_name if user_session else 'Unknown User'}")
                return False
            
        except Exception as e:
            error_msg = f"Error sending timesheet email: {str(e)}"
            self.logger.error(error_msg)
            import traceback
            self.logger.error(f"Traceback: {traceback.format_exc()}")
            return False
    
    def _create_safe_filename(self, original_filename, user_session, month_name, year):
        """Create a safe ASCII-only filename for email attachment"""
        try:
            # Extract timestamp from original filename
            timestamp_part = ""
            if "_MERQ_TIMESHEET_" in original_filename:
                timestamp_part = original_filename.split("_MERQ_TIMESHEET_")[-1]
            else:
                timestamp_part = datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            
            # Create safe name using English month names
            month_english_map = {
                'መስከረም': 'Meskerem', 'ጥቅምት': 'Tikimt', 'ኅዳር': 'Hidar', 
                'ታኅሣሥ': 'Tahsas', 'ጥር': 'Tir', 'የካቲት': 'Yekatit',
                'መጋቢት': 'Megabit', 'ሚያዝያ': 'Miyazya', 'ግንቦት': 'Ginbot',
                'ሰኔ': 'Sene', 'ሐምሌ': 'Hamle', 'ነሐሴ': 'Nehase', 'ጳጉሜ': 'Pagume'
            }
            
            english_month = month_english_map.get(month_name, month_name)
            
            # Clean user name for filename
            if user_session and user_session.full_name:
                clean_name = "".join(c for c in user_session.full_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                clean_name = clean_name.replace(' ', '_')
            else:
                clean_name = "Unknown_User"
            
            # Create safe filename
            safe_filename = f"{clean_name}_{english_month}_{year}_MERQ_TIMESHEET_{timestamp_part}"
            
            return safe_filename
            
        except Exception as e:
            self.logger.error(f"Error creating safe filename: {e}")
            # Fallback to simple filename
            return f"MERQ_Timesheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    def _create_email_body(self, user_session, month_name, year, hr_users, timesheet_file_path):
        """Create email body content with HTML formatting"""
        
        hr_names = [hr.get('full_name', 'HR Department') for hr in hr_users]
        hr_recipients = ", ".join(hr_names)
        filename = os.path.basename(timesheet_file_path)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        user_name = user_session.full_name if user_session else "Unknown User"
        user_position = user_session.position if user_session else "Unknown Position"
        user_department = user_session.department if user_session else "Unknown Department"
        user_employee_id = user_session.employee_id if user_session else "Unknown ID"
        user_supervisor = user_session.supervisor_name if user_session else "Unknown Supervisor"
        user_supervisor_position = user_session.supervisor_position_title if user_session else "Unknown Supervisor Position"
        
        # HTML email body with better formatting
        body = f"""
<html>
<head>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
        .header {{ background-color: #2C3E50; color: white; padding: 20px; text-align: center; }}
        .content {{ padding: 20px; }}
        .footer {{ background-color: #f8f9fa; padding: 15px; text-align: center; font-size: 12px; color: #666; }}
        .details {{ background-color: #f8f9fa; padding: 15px; border-left: 4px solid #3498DB; margin: 10px 0; }}
        .signature {{ margin-top: 20px; padding-top: 20px; border-top: 1px solid #ddd; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>MERQ CONSULTANCY PLC</h2>
        <h3>Timesheet Submission</h3>
    </div>
    
    <div class="content">
        <p>Dear <strong>{hr_recipients}</strong>,</p>
        
        <p>Please find attached the timesheet for your review and approval.</p>
        
        <div class="details">
            <h4>Employee Details:</h4>
            <ul>
                <li><strong>Name:</strong> {user_name}</li>
                <li><strong>Position:</strong> {user_position}</li>
                <li><strong>Department:</strong> {user_department}</li>
                <li><strong>Employee ID:</strong> {user_employee_id}</li>
                <li><strong>Supervisor:</strong> {user_supervisor}</li>
                <li><strong>Supervisor Position:</strong> {user_supervisor_position}</li>
            </ul>
        </div>
        
        <div class="details">
            <h4>Timesheet Information:</h4>
            <ul>
                <li><strong>Period:</strong> {month_name} {year}</li>
                <li><strong>Attachment:</strong> {filename}</li>
                <li><strong>Generated:</strong> {timestamp}</li>
            </ul>
        </div>
        
        <p>This timesheet has been generated and submitted through the MERQ Timesheet System.</p>
        <p>The attached Excel file contains the complete timesheet details formatted according to MERQ standards using the official template.</p>
        
        <div class="signature">
            <p>Best regards,</p>
            <p><strong>{user_name}</strong><br>
            {user_position}<br>
            MERQ Consultancy PLC</p>
        </div>
    </div>
    
    <div class="footer">
        <p>---<br>
        This is an automated email from MERQ Timesheet System.<br>
        <img src="https://merqconsultancy.org/wp-content/uploads/2023/06/merq.png" alt="MERQ Consultancy" width="100" height="30"><br>        
        MERQ Consultancy PLC | Excellence In Action!</p>
    </div>
</body>
</html>
"""
        return body
    
    def _send_email(self, msg, recipients):
        """Send the actual email"""
        try:
            self.logger.info(f"Connecting to SMTP server: {self.config['SMTPServer']}:{self.config['SMTPPort']}")
            server = smtplib.SMTP(self.config['SMTPServer'], self.config['SMTPPort'])
            self.logger.info("SMTP connection established")
            
            if self.config['UseTLS']:
                self.logger.info("Starting TLS")
                server.starttls()
                self.logger.info("TLS started")
            
            self.logger.info("Logging in to SMTP server")
            server.login(self.config['SMTPUser'], self.config['SMTPPassword'])
            self.logger.info("SMTP login successful")
            
            text = msg.as_string()
            self.logger.info(f"Sending email to {len(recipients)} recipients")
            server.sendmail(self.config['SMTPUser'], recipients, text)
            server.quit()
            self.logger.info("Email sent successfully")
            
            return True
            
        except Exception as e:
            self.logger.error(f"SMTP error: {str(e)}")
            import traceback
            self.logger.error(f"SMTP traceback: {traceback.format_exc()}")
            return False
    
    def test_connection(self):
        """Test SMTP connection"""
        try:
            server = smtplib.SMTP(self.config['SMTPServer'], self.config['SMTPPort'])
            
            if self.config['UseTLS']:
                server.starttls()
            
            server.login(self.config['SMTPUser'], self.config['SMTPPassword'])
            server.quit()
            
            self.logger.info("SMTP connection test successful")
            return True
            
        except Exception as e:
            self.logger.error(f"SMTP connection test failed: {str(e)}")
            return False

# Global instance
email_service = EmailService()