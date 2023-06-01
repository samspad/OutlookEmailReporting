#import sys
#sys.path.append("c:\users\shahidali\appdata\roaming\python\python310\site-packages")
import win32com.client as win32
import getpass
import csv

def get_time_by_subject(subject):
    if subject == 'Start and End Time for LCAP-ExportLoanPdsCompleted Normally':
        return '4:00 AM'
    elif subject == 'Start and End Time for BatchLoadSAORDTS_CFTFact_dailyCompleted Normally':
        return '7:00 AM'
    else:
        return ''

def read_outlook_emails(subject, csv_file):
    outlook_app = win32.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # Inbox folder

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # Sort emails by received time in descending order

    rows = []
    for item in items:
        if item.Subject == subject:
            #print("Subject:", item.Subject)
            #print("Received Time:", item.ReceivedTime)
            #print("Body:", item.Body)
            
            body_lines = item.Body.split("\n")
            milestone = body_lines[0].strip()
            start_time = body_lines[1].strip()
            end_time = body_lines[2].strip()
            time = get_time_by_subject(item.Subject)
            
            status = ""
            if end_time < time:
                status = "✔️ (Green)"
            else:
                status = "✔️ (Yellow)"
            
            rows.append([time, milestone, status, f"Start Time: {start_time}\nEnd Time: {end_time}"])

            with open(csv_file, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["Time", "Milestone", "Status", "Completed Time (Populate miss occurred)"])
                writer.writerows(rows)
            print(f"CSV file '{csv_file}' created successfully!")
            #break  # Assuming you only want to process the first matching email. If we remove this, it will look for all the emails with matching subject line

    outlook_app.Quit()

# Subject search examples
subject_to_search = [
    'Start and End Time for LCAP-ExportLoanPdsCompleted Normally',
    'Start and End Time for BatchLoadSAORDTS_CFTFact_dailyCompleted Normally'
]
output_csv_file = "email_data.csv"


# Prompt for Outlook credentials. This will be moved to .env file so that the credentials are hidden. We can
username = input("Enter your outlook username: ")
password = getpass.getpass("Enter your outlook password: ")

# Authenticate and read emails
outlook_app = win32.Dispatch("Outlook.Application")
namespace = outlook_app.GetNamespace("MAPI")
namespace.Logon(username, password, True, False)

read_outlook_emails(subject_to_search, output_csv_file)

# Log off
namespace.Logoff()
