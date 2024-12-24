import smtplib
import ssl
import time
import random
import base64
import os
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import threading
import queue
import uuid
from email.utils import formatdate, make_msgid, formataddr
import socket

class EmailAutomation:
    def __init__(self):
        self.smtp_servers = self.load_smtp_servers()
        
        # Configure logging
        logging.basicConfig(
            filename=f'email_logs_{datetime.now().strftime("%Y%m%d")}.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
    def load_smtp_servers(self):
        """Load SMTP servers"""
        servers = []
        with open('smtp_servers.txt', 'r') as f:
            for line in f:
                host, port, email, password = line.strip().split('|')
                servers.append({
                    'host': host,
                    'port': int(port),
                    'email': email,
                    'password': password
                })
        return servers
    
    def create_email(self, sender: str, recipient: str, subject: str, body: str) -> MIMEMultipart:
        """Create email with strict Outlook/Hotmail compatibility"""
        msg = MIMEMultipart('mixed')
        msg_alt = MIMEMultipart('alternative')
        
        # Sender information
        sender_domain = sender.split('@')[1]
        sender_name = sender.split('@')[0].replace('.', ' ').title()
        # sender_name = "Your Sender Name"
        message_id = f"<{uuid.uuid4().hex}@{sender_domain}>"
        
        # Essential headers with strict Outlook format
        msg['From'] = f'"{sender_name}" <{sender}>'
        msg['To'] = recipient
        msg['Subject'] = subject
        msg['Date'] = formatdate(localtime=True)
        msg['Message-ID'] = message_id
        
        # Strict Outlook/Exchange headers
        msg['X-MS-Exchange-Organization-MessageDirectionality'] = 'Originating'
        msg['X-MS-Exchange-Organization-AuthSource'] = sender_domain
        msg['X-MS-Exchange-Organization-AuthAs'] = 'Internal'
        msg['X-MS-Exchange-Organization-AuthMechanism'] = '05'
        msg['X-MS-Exchange-Organization-AVStamp-Enterprise'] = '1.0'
        msg['X-MS-Exchange-Organization-SCL'] = '-1'
        
        # Additional Exchange headers
        msg['X-MS-Exchange-CrossTenant-originalarrivaltime'] = formatdate(localtime=True)
        msg['X-MS-Exchange-CrossTenant-fromentityheader'] = 'Hosted'
        msg['X-MS-Exchange-Transport-CrossTenantHeadersStamped'] = '1'
        msg['X-MS-Exchange-CrossTenant-id'] = uuid.uuid4().hex
        
        # Microsoft specific headers
        msg['X-Mailer'] = 'Microsoft Outlook 16.0'
        msg['X-Microsoft-Antispam'] = 'BCL:0;'
        msg['X-Microsoft-Antispam-Message-Info'] = f'=?utf-8?B?{base64.b64encode(message_id.encode()).decode()}?='
        msg['X-Microsoft-Exchange-Diagnostics'] = uuid.uuid4().hex
        
        # Create plain text version
        body_text = body.replace('<p>', '').replace('</p>', '\n').replace('<br>', '\n').replace('<li>', '- ').replace('</li>', '\n')
        plain_text = f"""
                    {subject}

                    {body_text}

                    Best regards,
                    {sender_name}
                """.strip()
                
        text_part = MIMEText(plain_text, 'plain', 'utf-8')
        msg_alt.attach(text_part)
        
        # Create HTML with strict Outlook formatting
        html_content = f'''
        <!DOCTYPE html>
        <html>
            <head>
                <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
            </head>
            <body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000;">
                <div style="max-width: 600px; margin: 0 auto;">
                    {body}
                </div>
            </body>
        </html>
        '''
        html_part = MIMEText(html_content, 'html', 'utf-8')
        msg_alt.attach(html_part)
        
        msg.attach(msg_alt)
        return msg
    
    def send_single_email(self, recipient: str, subject: str, body: str) -> bool:
        """Send a single email using the provided SMTP server"""
        server = self.smtp_servers[1]  # Using the first (only) SMTP server
        success = True
        try:
            with smtplib.SMTP(server['host'], server['port'], timeout=30) as smtp:
                # Initial EHLO with full domain
                smtp.ehlo_or_helo_if_needed()
                
                # TLS setup
                context = ssl.create_default_context()
                smtp.starttls(context=context)
                
                # Second EHLO after TLS
                smtp.ehlo_or_helo_if_needed()
                
                # Login
                smtp.login(server['email'], server['password'])
                
                msg = self.create_email(
                    sender=server['email'],
                    recipient=recipient,
                    subject=subject,
                    body=body
                )
                
                smtp.sendmail(server['email'], recipient, msg.as_string())
                logging.info(f"Email sent successfully to {recipient}")
                print(f"Email sent successfully to {recipient}")
                return True
                
        except Exception as e:
            logging.error(f"Error sending email to {recipient}: {str(e)}")
            print(f"Error sending email to {recipient}: {str(e)}")
            return False
    
    def send_emails(self, recipients: list[str], subject: str, body: str) -> bool:
        """Send emails to multiple recipients using a single SMTP connection"""
        server = self.smtp_servers[1]  # Using the first (only) SMTP server
        
        try:
            with smtplib.SMTP(server['host'], server['port'], timeout=30) as smtp:
                # Initial EHLO with full domain
                smtp.ehlo_or_helo_if_needed()
                print("Connected to SMTP server")
                
                # TLS setup
                context = ssl.create_default_context()
                smtp.starttls(context=context)
                print("STARTTLS completed")
                
                # Second EHLO after TLS
                smtp.ehlo_or_helo_if_needed()
                
                # Login
                smtp.login(server['email'], server['password'])
                print("Login successful")
                
                for i,recipient in enumerate(recipients):
                    try:
                        msg = self.create_email(
                            sender=server['email'],
                            recipient=recipient,
                            subject=subject,
                            body=body
                        )
                        
                        smtp.sendmail(server['email'], recipient, msg.as_string())
                        logging.info(f"Email sent successfully to {recipient}")
                        print(f"({i+1}/{len(recipients)}) Email sent successfully to {recipient} ")
                        time.sleep(random.randint(5, 10))  # Delay between sends
                        
                    except Exception as e:
                        logging.error(f"Error sending email to {recipient}: {str(e)}")
                        print(f"({i+1}/{len(recipients)}) Error sending email to {recipient} : {str(e)}")
                        time.sleep(random.randint(5, 10))  # Delay between sends
        
            print(f"Script finished successfully")     
            return True
                    
        except Exception as e:
            logging.error(f"Error establishing SMTP connection: {str(e)}")
            print(f"Error establishing SMTP connection: {str(e)}")
            return False








if __name__ == "__main__":
    automation = EmailAutomation()
    with open('template.html', 'r') as f:
        template = f.read()
        
    with open('data/recipients.txt', 'r') as f:
        recipients = f.readlines()
        
    automation.send_single_email(
        recipient="yourEmail@example.com",
        subject="Your Subject",
        body=template,
    )

    # automation.send_emails(
    #     recipients=recipients,
    #     subject="Your Subject",",
    #     body=template,
    # )
