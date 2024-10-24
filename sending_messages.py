# -*- coding: utf-8 -*-
"""
This file contains functions to send messages via email, Teams, and SMS

Created on Wed Oct 16 13:29:29 2024

@author: HADIJK
"""

import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
import json

# This file contains 

# Function to send email
def send_email_with_html(subject, body, email_receiver, sMeetsetFolder, previous_day):
    email_sender ="monitoringheatpump@gmail.com"
    email_password = "crhr pwiw pwky icsf"

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Heatpump Monitoring"
    msg['From'] = email_sender
    msg['To'] = email_receiver

    # html = '<html><body><p>Hi, I have the following alerts for you!</p></body></html>'
    intro_message = f"Hello colleagues,\n\nHeatpump monitoring saw a problem in the daily data for {sMeetsetFolder} at {previous_day}.\n"
    end_message = "\n\nGreetings,\nHeatpump Monitoring System"
    full_body = intro_message + body + end_message

    # Create the HTML part of the email
    html = f'<html><body><p>{full_body.replace("\n", "<br>")}</p></body></html>'
    part2 = MIMEText(html, 'html')

    msg.attach(part2)

    # Send the message via gmail's regular server, over SSL - passwords are being sent, afterall
    s = smtplib.SMTP_SSL('smtp.gmail.com')
    # uncomment if interested in the actual smtp conversation
    # s.set_debuglevel(1)
    # do the smtp auth; sends ehlo if it hasn't been sent already
    s.login(email_sender, email_password)

    s.sendmail(email_sender, email_receiver, msg.as_string())
    s.quit()


def send_email(subject, body, email_receiver, sMeetsetFolder, previous_day):
    email_sender = "monitoringheatpump@gmail.com"
    email_password = "crhr pwiw pwky icsf"  # Replace with your app password if 2FA is enabled

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Heatpump Monitoring"
    msg['From'] = email_sender
    msg['To'] = email_receiver

    intro_message = f"Hello colleagues,\n\nDaily data report for {sMeetsetFolder} at {previous_day}.\n"
    end_message = "\n\nGreetings,\nHeatpump Monitoring System"
    full_body = intro_message + body + end_message

    # Attach the plain text part to the email
    part1 = MIMEText(full_body, 'plain')
    msg.attach(part1)

    try:
        # Send the message via Gmail's regular server, using STARTTLS
        s = smtplib.SMTP('smtp.gmail.com', 587)
        # s.set_debuglevel(1)  # Enable debug output
        s.ehlo()  # Identify ourselves to the SMTP server
        s.starttls()  # Secure the connection
        s.ehlo()  # Re-identify ourselves as an encrypted connection
        s.login(email_sender, email_password)
        s.sendmail(email_sender, email_receiver, msg.as_string())
        s.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")

def send_teams_message(body, sMeetsetFolder, previous_day):
    webhook_url = "https://dnv.webhook.office.com/webhookb2/de5c61a7-826f-4f87-9c9f-93f5366aa625@adf10e2b-b6e9-41d6-be2f-c12bb566019c/IncomingWebhook/6dd787245df144fba6398bbdd59c473a/5fcef47d-1ed1-4d15-92b3-dc1169d4a35e/V2fqnkFCPWLIE4NfhKdZuU3tUpztHS4FFeH743E_yqXTY1"
    intro_message = f"Daily data report for {sMeetsetFolder} at {previous_day}.\n\n"
    full_message = intro_message + body

    headers = {'Content-Type': 'application/json'}
    payload = {
        "text": full_message
    }
    
    response = requests.post(webhook_url, headers=headers, data=json.dumps(payload))
    
    if response.status_code == 200:
        print("Message sent to Teams successfully.")
    else:
        print(f"Failed to send message to Teams: {response.status_code}, {response.text}")


def send_sms_via_email(phone_number, carrier_domain, subject, message, email_sender, email_password):
    # Construct the email
    sms_recipient = f"{phone_number}@{carrier_domain}"
    msg = MIMEText(message)
    msg['From'] = email_sender
    msg['To'] = sms_recipient
    msg['Subject'] = subject

    try:
        # Send the message via Gmail's SMTP server
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(email_sender, email_password)
        s.sendmail(email_sender, sms_recipient, msg.as_string())
        s.quit()
        print("SMS sent successfully via email")
    except Exception as e:
        print(f"Failed to send SMS: {e}")