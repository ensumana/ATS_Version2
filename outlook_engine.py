import os
import re
import win32com.client
import pythoncom



RF_KEYWORDS = [
    "rf testing",
    "rf design",
    "rf architect",
    "engineer",
    "resume"
]

MANDATORY_WORD = "application"


def clean(text):
    if not text:
        return ""
    return re.sub(r"\s+", " ", str(text).lower())


def get_sender_name(mail):
    try:
        if mail.SenderEmailType == "EX":
            user = mail.Sender.GetExchangeUser()
            if user:
                return user.Name
    except:
        pass

    return mail.SenderEmailAddress.split("@")[0]


def is_matching(subject, body):
    text = clean(subject + " " + body)

    if MANDATORY_WORD not in text:
        return False

    return any(k in text for k in RF_KEYWORDS)


def clean_filename(name):
    """
    Remove invalid Windows filename characters
    """
    return re.sub(r'[<>:"/\\|?*]', '_', str(name))


def download_resumes(folder_path,
                     from_date,
                     to_date,
                     progress_callback=None):

    processed = 0
    downloaded = 0

    pythoncom.CoInitialize()

    try:

        outlook = win32com.client.gencache.EnsureDispatch(
            "Outlook.Application"
        ).GetNamespace("MAPI")

        inbox = outlook.GetDefaultFolder(6)

        messages = inbox.Items

        messages.Sort("[ReceivedTime]", True)

        total_messages = messages.Count

        print(f"Total emails in Inbox: {total_messages}")

        for idx, message in enumerate(messages, start=1):

            try:

                if not hasattr(message, "ReceivedTime"):
                    continue

                received_time = message.ReceivedTime

                # Convert Outlook COM datetime
                received_time = received_time.replace(tzinfo=None)

                # Date filtering
                if received_time < from_date:
                    continue

                if received_time > to_date:
                    continue

                processed += 1

                sender_name = clean_filename(
                    str(message.SenderName).strip()
                )

                attachments = message.Attachments

                for i in range(1, attachments.Count + 1):

                    attachment = attachments.Item(i)

                    filename = attachment.FileName

                    ext = os.path.splitext(filename)[1].lower()

                    if ext not in [".pdf", ".doc", ".docx"]:
                        continue

                    save_name = f"{sender_name}{ext}"

                    save_path = os.path.join(
                        folder_path,
                        save_name
                    )

                    counter = 1

                    while os.path.exists(save_path):

                        save_name = (
                            f"{sender_name}_{counter}{ext}"
                        )

                        save_path = os.path.join(
                            folder_path,
                            save_name
                        )

                        counter += 1

                    attachment.SaveAsFile(save_path)

                    downloaded += 1

                    print(f"Saved: {save_name}")

                if progress_callback:

                    progress = int(
                        (idx / max(total_messages, 1)) * 100
                    )

                    progress_callback(progress)

            except Exception as email_error:

                print(
                    f"Error processing email: "
                    f"{email_error}"
                )

                continue

        return processed, downloaded

    finally:

        pythoncom.CoUninitialize()

from datetime import timezone

def make_naive(dt):
    """Convert Outlook timezone-aware datetime → naive"""
    if dt.tzinfo is not None:
        return dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt

