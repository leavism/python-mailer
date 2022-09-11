from email.mime.base import MIMEBase
import smtplib
import ssl
import csv
import logging
from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

import config

class Mailer:
  """Auto-emailer for Gmail."""

  def __init__(self):
    self.port = 465
    self.smtp_server_domain_name = "smtp.gmail.com"
    self.sender_mail = config.login["email"]
    self.password = config.login["password"]

    logging.basicConfig(
                        format="%(levelname)s: %(message)s",
                        level=20,
                        handlers = [
                          logging.FileHandler("emailing.log"),
                          logging.StreamHandler()
                        ]
    )

  def load_emails(self):
    """Loads the email from a CSV file. Returns an array of dictionaries."""
    mailing_list = []
    with open("mailing-list.csv") as csvfile:
      reader = csv.DictReader(csvfile)
      logging.info("Loaded emails")
      for contact in reader:
        mailing_list.append(contact)
    return mailing_list

  def filter_emails(self, mailing_list):
    """Filters through the mailing list and calls #send on contacts that have an email address. Skips any contacts that doesn't have an email address. Logs everything in the email.log."""
    for contact in mailing_list:
      logging.info("Looking at %s" %(contact["Company"]))
      if contact["Email"]:
        try:
          logging.info("Emailing {0} at address {1}".format((contact["Contact Name"] or "no name"), contact["Email"]))
          self.send(contact)
          sleep(5)
        except:
          logging.exception("Uh oh! Something went wrong")
      else:
        logging.warning("No email for {0}- skipping {1}".format(contact["Company"], contact["Contact Name"]))

  def send(self, contact):
    """The actual emailing method."""
    ssl_context = ssl.create_default_context()
    service = smtplib.SMTP_SSL(self.smtp_server_domain_name, self.port, context=ssl_context)
    service.login(self.sender_mail, self.password)
    
    # Sets subject, from, and to for the email
    mail = MIMEMultipart('alternative')
    mail['Subject'] = 'SF Hacks 2021-2022 Partnership Opportunity'
    mail['From'] = self.sender_mail
    mail['To'] = contact["Email"]

    # Loads html template
    html_template = ""
    with open("2022_email_draft.html", "r") as file:
      html_template = file.read()

    # Interpolate inputs into the html template
    html_template = html_template.format(config.login["sender-name"], config.login["sender-position"], config.login["sender-fullname"])

    # Writes the html template to email body
    html_content = MIMEText(html_template, "html")
    mail.attach(html_content)

    file_path = "2022_sponsorship_brochure.pdf"
    mimeBase = MIMEBase("application", "octet-stream")
    with open(file_path, "rb") as file:
      mimeBase.set_payload(file.read())
    encoders.encode_base64(mimeBase)
    mimeBase.add_header("Content-Disposition", f"attachment; filename={Path(file_path).name}")
    mail.attach(mimeBase)
    # Sends mail
    service.sendmail(self.sender_mail, contact["Email"], mail.as_string())

    service.quit()

if __name__ == '__main__':
    mail = Mailer()
    mailing_list = mail.load_emails()
    mail.filter_emails(mailing_list)
