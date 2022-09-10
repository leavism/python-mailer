import smtplib
import ssl
import csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import config

class Mailer:
  def __init__(self):
    self.port = 465
    self.smtp_server_domain_name = "smtp.gmail.com"
    self.sender_mail = config.login["email"]
    self.password = config.login["password"]

  def load_emails(self):
    mailing_list = []
    with open("mailing-list.csv") as csvfile:
      reader = csv.DictReader(csvfile)
      for contact in reader:
        mailing_list.append(contact)
    return mailing_list

  def send(self, emails):
    ssl_context = ssl.create_default_context()
    service = smtplib.SMTP_SSL(self.smtp_server_domain_name, self.port, context=ssl_context)
    service.login(self.sender_mail, self.password)
    
    for email in emails:
      mail = MIMEMultipart('alternative')
      mail['Subject'] = 'Testing Emails'
      mail['From'] = self.sender_mail
      mail['To'] = email

      text_template = """
      Hi {0},
      What goes after Hi?
      """

      text_content = MIMEText(text_template.format(email.split("@")[0]), 'plain')
      mail.attach(text_content)

      service.sendmail(self.sender_mail, email, mail.as_string())

    service.quit()

if __name__ == '__main__':
    # mails = input("Enter emails: ").split()

    mail = Mailer()
    # mail.send(mails)
    emails = mail.load_emails()
    for contact in emails:
      print(contact)
