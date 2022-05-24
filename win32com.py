import datetime
import os
import win32com.client

path = r"D:\centos_2GB"
today = datetime.date.today()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items


def save_attachments(subject):
    for message in messages:
        if message.Subject == subject:
        # if message.Subject == subject and message.Unread or message.Senton.date() == today:
            for attachment in message.Attachments:
                print(attachment.FileName)
                attachment.SaveAsFile(os.path.join(path, str(attachment)))


if __name__ == "__main__":
    save_attachments('batch files')