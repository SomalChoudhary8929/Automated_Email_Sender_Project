import smtplib
import os
import json
import logging
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import schedule
import time
from datetime import datetime
from typing import List, Optional
import re

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_automation.log'),
        logging.StreamHandler()
    ]
)

class EmailAutomation:
    def __init__(self, config_file: str = 'config.json'):
        self.config_file = config_file
        self.config = self._load_config()
        self.is_running = True

    def _load_config(self) -> dict:
        """Load configuration from JSON file"""
        if not os.path.exists(self.config_file):
            default_config = {
                "email": {
                    "smtp_server": "smtp.gmail.com",
                    "smtp_port": 465,
                    "sender_email": "",
                    "app_password": ""
                },
                "schedule": {
                    "times": ["19:03"],
                    "days": ["monday", "tuesday", "wednesday", "thursday", "friday"]
                },
                "recipients": [],
                "email_template": {
                    "subject": "Automated Email",
                    "body": "Hello! This is an automated email sent by somal Choudhary."
                }
            }
            with open(self.config_file, 'w') as f:
                json.dump(default_config, f, indent=4)
            logging.info(f"Created default config file: {self.config_file}")
            return default_config
        
        with open(self.config_file, 'r') as f:
            return json.load(f)

    def validate_email(self, email: str) -> bool:
        """Validate email address format"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))

    def send_email(self, recipient: str, subject: Optional[str] = None, 
                  body: Optional[str] = None, attachments: List[str] = None) -> bool:
        """Send email to a single recipient with optional attachments"""
        try:
            if not self.validate_email(recipient):
                logging.error(f"Invalid email address: {recipient}")
                return False

            msg = MIMEMultipart()
            msg["Subject"] = subject or self.config["email_template"]["subject"]
            msg["From"] = self.config["email"]["sender_email"]
            msg["To"] = recipient
            msg.attach(MIMEText(body or self.config["email_template"]["body"], "plain"))

            # Add attachments if any
            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as f:
                            part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
                            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                            msg.attach(part)
                    else:
                        logging.warning(f"Attachment not found: {file_path}")

            with smtplib.SMTP_SSL(
                self.config["email"]["smtp_server"],
                self.config["email"]["smtp_port"]
            ) as smtp:
                smtp.login(
                    self.config["email"]["sender_email"],
                    self.config["email"]["app_password"]
                )
                smtp.send_message(msg)

            logging.info(f"Email sent successfully to {recipient}")
            return True

        except Exception as e:
            logging.error(f"Failed to send email to {recipient}: {str(e)}")
            return False

    def send_bulk_emails(self) -> None:
        """Send emails to all recipients in the configuration"""
        for recipient in self.config["recipients"]:
            self.send_email(recipient)

    def start(self) -> None:
        """Start the email automation"""
        logging.info("Email automation started... Press Ctrl+C to stop.")
        
        # Schedule emails for each configured time
        for time_str in self.config["schedule"]["times"]:
            schedule.every().day.at(time_str).do(self.send_bulk_emails)

        while self.is_running:
            try:
                schedule.run_pending()
                time.sleep(5)
            except KeyboardInterrupt:
                self.stop()
                break

    def stop(self) -> None:
        """Stop the email automation"""
        self.is_running = False
        logging.info("Email automation stopped.")

def main():
    automation = EmailAutomation()
    
    # Check if email credentials are configured
    if not automation.config["email"]["sender_email"] or not automation.config["email"]["app_password"]:
        print("Please configure your email credentials in config.json")
        return

    # Check if recipients are configured
    if not automation.config["recipients"]:
        print("Please add recipients to config.json")
        return

    automation.start()

if __name__ == "__main__":
    main()