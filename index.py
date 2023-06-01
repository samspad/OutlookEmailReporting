import sys
#sys.path.append("c:\users\shahidali\appdata\roaming\python\python310\site-packages")
import win32com.client as win32
import getpass
import csv

def read_outlook_emails(subject, csv_file):
    outlook_app = win32.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # Inbox folder

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # Sort emails by received time in descending order

    for item in items:
        if item.Subject == subject:
            #print("Subject:", item.Subject)
            #print("Received Time:", item.ReceivedTime)
            #print("Body:", item.Body)
            
            body_lines = item.Body.split("\n")
            rows = []
            for i in range(0, len(body_lines), 3):
                milestone = body_lines[i].strip()
                time = body_lines[i+1].strip()
                completed_time = body_lines[i+2].strip()
                rows.append([time, milestone, completed_time])

            with open(csv_file, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["Time", "Milestone", "Completed Time"])
                writer.writerows(rows)
            print(f"CSV file '{csv_file}' created successfully!")
            break  # Assuming you only want to process the first matching email. If we remove this, it will look for all the emails with matching subject line

    outlook_app.Quit()

# Example usage
subject_to_search = "Start and End Time for LCAP-ExportLoanPdsCompleted Normally"
output_csv_file = "email_data.csv"


# Prompt for Outlook credentials. This will be moved to .env file so that the credentials are hidden. We can
username = input("Enter your Outlook username: ")
password = getpass.getpass("Enter your Outlook password: ")

# Authenticate and read emails
outlook_app = win32.Dispatch("Outlook.Application")
namespace = outlook_app.GetNamespace("MAPI")
namespace.Logon(username, password, True, False)

read_outlook_emails(subject_to_search, output_csv_file)

# Log off (this will close the outlook application)
#namespace.Logoff()
