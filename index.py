import sys
#sys.path.append("c:\users\shahidali\appdata\roaming\python\python310\site-packages")

import win32com.client as win32

import getpass

def read_outlook_emails(subject):
    outlook_app = win32.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # Inbox folder

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # Sort emails by received time in descending order

    for item in items:
        if item.Subject == subject:
            print("Subject:", item.Subject)
            print("Received Time:", item.ReceivedTime)
            print("Body:", item.Body)
            break  # Assuming you only want to retrieve the first matching email

    outlook_app.Quit()

# Example usage
subject_to_search = "Start and End Time for LCAP-ExportLoanPdsCompleted Normally"

# Prompt for Outlook credentials
username = input("Enter your Outlook username: ")
password = getpass.getpass("Enter your Outlook password: ")

# Authenticate and read emails
outlook_app = win32.Dispatch("Outlook.Application")
namespace = outlook_app.GetNamespace("MAPI")
namespace.Logon(username, password, True, False)

read_outlook_emails(subject_to_search)

# Log off
namespace.Logoff()
