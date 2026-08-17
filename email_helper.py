import smtplib
import json
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

CONFIG_FILE = "email_config.json"

def load_email_config():
    """
    Loads email configuration from a local JSON file.
    """
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print("Error loading email config:", e)
    return {}

def save_email_config(config):
    """
    Saves email configuration to a local JSON file.
    """
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print("Error saving email config:", e)
        return False

def send_email(smtp_server, smtp_port, sender_email, sender_password, recipient_email, subject, body):
    """
    Sends an email using standard SMTP.
    """
    try:
        # Create message container
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        # Attach the body text
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Connect to server using TLS
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()  # Upgrade connection to secure TLS
        server.login(sender_email, sender_password)
        
        # Send mail
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        return True, "Correo enviado exitosamente."
    except Exception as e:
        return False, f"Error al enviar correo: {str(e)}"
