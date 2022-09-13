import smtplib
import ssl
import csv
import logging
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path


import config

class Mailer:
  """Auto-emailer for Gmail. Wrote and tested this on Python 3.8.9. All libraries used in this script are built-in modules and do not need to be installed separately."""

  def __init__(self):
    # Don't change these
    self.port = 465
    self.smtp_server_domain_name = "smtp.gmail.com"

    # Login info to sender gmail account. Edit these from the config file.
    self.sender_address = config.login["email"]
    self.password = config.login["password"]

    # File paths used in auto-emailer. Edit these from the config file.
    self.brochure_file_path = config.info["brochure_file_path"]
    self.html_template_file_path = config.info["html_template_file_path"]
    self.mailing_list_file_path = config.info["mailing_list_file_path"]

    # Additional info for the email. Edit these from the config file.
    self.full_name = config.sender["full_name"]
    self.officer_position = config.sender["officer_position"]
    self.email_subject = config.info["email_subject"]

    # Logging config
    self.currentDT = datetime.datetime.now()
    logging.basicConfig(
                        format="%(levelname)s: %(message)s",
                        level=10,
                        handlers = [
                          logging.FileHandler("{0}.log".format(self.currentDT.strftime("%m-%d %H%M%S"))),
                          logging.StreamHandler()
                        ]
    )

  def load_emails(self):
    """Loads the email from a CSV file. Returns a list of dictionaries containing info from the CSV."""
    try:
      with open(self.mailing_list_file_path) as csvfile:
        reader = csv.DictReader(csvfile)
        logging.info("Loaded emails")
        contact_list = []
        for contact in reader:
          contact_list.append(contact)
    except:
      logging.exception("Uh oh! Couldn't load contacts from the email list.")
    return contact_list

  def filter_emails(self, mailing_list):
    no_email = []
    others = []
    dup_contacts = []
    filtered_contacts = []

    to_send = []

    for contact in mailing_list:
      if contact["Email"] in filtered_contacts:
        logging.warning("%s may be a duplicate contact." %(contact))
        dup_contacts.append(contact)
        continue

      if contact["Email"].count("@") == 0 and contact["Email"].count(".com") >= 1:
        """Checks for possible linkedin links rather than emails"""
        logging.warning("{0} might contain a link and not an email-- SKIPPING".format(contact))
        others.append(contact)
      elif contact["Email"].count("@") not in [0, 1]:
        """Checks for multiple emails in the same cell"""
        logging.warning("{0} something is weird about this contact-- SKIPPING".format(contact))
        others.append(contact)
      elif not contact["Email"]:
        """Checks for contacts with no email addresses"""
        logging.warning("No email address for company {0}- SKIPPING".format(contact["Company"]))
        no_email.append(contact)
      else:
        logging.info("Contact {0} ({2}) at address {1} seems correct. Added to final mailing list.".format((contact["Contact Name"] or "no name"), contact["Email"], contact["Company"] or "no company"))
        filtered_contacts.append(contact["Email"])
        to_send.append(contact)

    logging.debug("=====================\n")
    logging.warning("Not emailing (duplicates) (size: {1}):\n{0}".format(self.list_to_string(dup_contacts), len(dup_contacts)))
    logging.warning("Not emailing (no email address) (size: {1}):\n{0}".format(self.list_to_string((no_email)), len(no_email)))
    logging.warning("Not emailing (unknown reason) (size {1}):\n{0}".format(self.list_to_string((others)), len(others)))
    logging.info("Emailing (size: {1}):\n{0}".format(self.list_to_string(to_send), len(to_send)))

    return {"no_email_address": no_email, "duplicate_contacts": dup_contacts, "to_send": to_send}

  def list_to_string(self, mailing_list):
    string = ""
    for contact in mailing_list:
      string += "Company: {0}\nContact Name: {1}\nEmail: {2}\n-----\n".format(contact["Company"], contact["Contact Name"], contact["Email"])
    return string

  def send(self, contact):
    """The actual emailing method."""
    ssl_context = ssl.create_default_context()
    service = smtplib.SMTP_SSL(self.smtp_server_domain_name, self.port, context=ssl_context)
    service.login(self.sender_address, self.password)
    
    # Sets subject, from, and to for the email
    mail = MIMEMultipart('alternative')
    mail['Subject'] = self.email_subject
    mail['From'] = self.sender_address
    mail['To'] = contact["Email"]

    # Loads html template
    try:
      with open(self.html_template_file_path, "r") as file:
        html_template = file.read()
    except:
      logging.exception("Couldn't load the html template")

    # Interpolate inputs into the html template
    html_template = html_template.format(
                                        self.full_name.split(" ")[0], 
                                        self.officer_position, 
                                        self.full_name
                                        )

    # Writes the html template to email body
    html_content = MIMEText(html_template, "html")
    mail.attach(html_content)

    # Attaches the .pdf brochure to email
    mimeBase = MIMEBase("application", "octet-stream")
    try:
      with open(self.brochure_file_path, "rb") as file:
        mimeBase.set_payload(file.read())
      encoders.encode_base64(mimeBase)
      mimeBase.add_header("Content-Disposition", f"attachment; filename={Path(self.brochure_file_path).name}")
      mail.attach(mimeBase)
    except:
      logging.exception("Uh oh! Couldn't load the .pdf brochure")

    # Sends mail
    service.sendmail(self.sender_address, contact["Email"], mail.as_string())

    # Closes the connection to gmail
    service.quit()

if __name__ == '__main__':
    mail = Mailer()
    mailing_list = mail.load_emails()
    mail.filter_emails(mailing_list)

    print("\n\nCheck the latest {0}.log to make sure these are the contacts you want to email.\n".format(mail.currentDT.strftime("%m-%d %H%M%S")) +
    "Type \"Y\" to send emails to the contacts listed in the log-- skipping the duplicates and contacts without address." +
    "\nType \"N\" to cancel the operation entirely.")
