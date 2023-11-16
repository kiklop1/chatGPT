import openai
import pandas as pd
import time
import re
import win32com.client
import openpyxl


def send_outlook_email(subject, body, to_email):
  outlook_app = win32com.client.Dispatch("Outlook.Application")
  namespace = outlook_app.GetNamespace("MAPI")

  # Specify the account for sending the email
  accounts = namespace.Accounts
  for account in accounts:
    if account.SmtpAddress == "financiallyfitinvestor@financiallyfitinvestor.com":  # Replace with your email address
      sender_account = account
      break
  else:
    raise Exception("Email account not found")

  # Create a new mail item
  mail_item = outlook_app.CreateItem(0)  # 0 represents olMailItem (Mail item)

  # Set the email properties
  mail_item.Subject = subject
  mail_item.Body = body
  mail_item.To = to_email

  # Set the account for sending the email
  mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, sender_account))

  # Send the email
  mail_item.Send()

api_key = "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"
openai.api_key = api_key

subject_pattern = re.compile(r"Subject: (.+?)\n", re.DOTALL)

list_urlova = pd.read_excel(r"C:\Users\BzNs\PycharmProjects\Startup\startup_database_test v2 - Copy.xlsx")
filtered_data = list_urlova[list_urlova['Founded'] == 2022]

outlook_account = None
for account in win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").Accounts:
    if account.SmtpAddress == "financiallyfitinvestor@financiallyfitinvestor.com":
        outlook_account = account
        break

for index, row in filtered_data.iterrows():
    email, sent, business_desc, startup = row['Email'], row['Sent'], row['Business Description'], row['Startup']

    if pd.notnull(email) and pd.isnull(sent):
      print(startup)
      client = openai.ChatCompletion.create(
          model="gpt-3.5-turbo",
          messages=[
              {"role": "system", "content": "You are a marketing expert, skilled in constructing personalized email content."},
              {"role": "user", "content": f"Compose personalized content with mild words and many conjunctions without an introduction and ending. Contents need to be in size of 150-200 words. Here are the details about the startup toward which the email will be sent:\n"
                                           f"startup name to which the email will be sent: {startup}\n"
                                           f"business description to which the email will be sent: {business_desc}\n"
                                           "Here are the details about me and my company: My name: Kiklop, Company: Cyclop Corp, description: Cyclop Corp specializes in financial advisory to clients with a focus on startups, creating valuations, financial projections, business plans, pitch desks, etc"
               }
          ]
      )

      subject_match = subject_pattern.search(client.choices[0].message["content"])

      if subject_match:
        print(subject_match)
        subject = subject_match.group(1)
        # Example usage
        subject_ = "Maximize Your Finances: Explore Projections, Valuations, and ESG Reporting Support"
        body_ = subject_match
        recipient_ = 'matejpretkovic4@gmail.com'

        send_outlook_email(subject_, body_, recipient_)

        wb = openpyxl.load_workbook(r"C:\Users\BzNs\PycharmProjects\Startup\startup_database_test v2 - Copy.xlsx")
        sheet = wb.active
        sheet.cell(row=index + 1, column=13, value="Sent")
        wb.save(r"C:\Users\BzNs\PycharmProjects\Startup\startup_database_test v2 Copy v3.xlsx")
