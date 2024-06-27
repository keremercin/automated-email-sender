import smtplib
import base64
import time
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

def send_emails(excel_file_path, html_file_path, image_file_path, smtp_server, smtp_port, username, password):
    # Read the Excel file
    df = pd.read_excel(excel_file_path)

    # Read the HTML file
    with open(html_file_path, 'r', encoding='latin-1') as file:
        mail_content = file.read()

    # Connect to the SMTP server
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)

    # Read the cover image
    with open(image_file_path, 'rb') as img_file:
        msg_image = MIMEImage(img_file.read())
        msg_image.add_header('Content-ID', '<cover_image>')  # Add an ID for referencing the image in HTML

    # Send email to each recipient
    for index, row in df.iterrows():
        recipient_name = row['full_name']
        recipient_email = row['email1']

        msg = MIMEMultipart()
        msg['From'] = username
        msg['To'] = recipient_email
        msg['Subject'] = "Pathograma Invitation"

        # Personalize the HTML content and add the image
        personalized_content = mail_content.replace('Dr Luis Sardina', f'Dr {recipient_name}')
        html_img_tag = '<img width=674 height=117 src="cid:cover_image">'
        full_html_content = html_img_tag + personalized_content
        msg.attach(MIMEText(full_html_content, 'html'))
        # Attach the image to the message
        msg.attach(msg_image)

        try:
            server.send_message(msg)
            print(f"Email sent to {recipient_name}.")
        except Exception as e:
            print(f"Error: {e}")

        time.sleep(30)

        if (index + 1) % 50 == 0:
            server.quit()
            time.sleep(10)
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(username, password)

    server.quit()

if __name__ == "__main__":
    excel_file_path = "path/to/your/excel_file.xlsx"  # Change this to the path of your Excel file
    html_file_path = "path/to/your/html_file.htm"  # Change this to the path of your HTML file
    image_file_path = "path/to/your/image_file.png"  # Change this to the path of your image file
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    username = "your_email@gmail.com"  # Change this to your email
    password = "your_password"  # Change this to your email password

    send_emails(excel_file_path, html_file_path, image_file_path, smtp_server, smtp_port, username, password)
