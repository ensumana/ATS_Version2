import os
import pythoncom
import win32com.client
from datetime import timezone
import json
from datetime import datetime


def normalize_datetime(dt):
    if not dt:
        return None
    try:
        if hasattr(dt, "tzinfo") and dt.tzinfo:
            return dt.astimezone(timezone.utc).replace(tzinfo=None)
    except:
        pass
    return dt


def get_outlook_namespace():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return outlook


def get_mailboxes():
    pythoncom.CoInitialize()
    try:
        outlook = get_outlook_namespace()
        return [outlook.Folders.Item(i).Name for i in range(1, outlook.Folders.Count + 1)]
    finally:
        pythoncom.CoUninitialize()


def download_resumes(
        folder_path,
        from_date,
        to_date,
        progress_callback=None,
        mailbox_name=None):
    resume_metadata = {}
    pythoncom.CoInitialize()

    processed = 0
    downloaded = 0

    try:
        outlook = get_outlook_namespace()

        # --------------------------------------------
        # Select mailbox
        # --------------------------------------------
        if mailbox_name:

            selected_store = None

            for i in range(1, outlook.Folders.Count + 1):

                store = outlook.Folders.Item(i)

                if store.Name == mailbox_name:
                    selected_store = store
                    break

            if selected_store is None:
                raise Exception(
                    f"Mailbox not found: {mailbox_name}"
                )

            print(f"Selected Mailbox : {selected_store.Name}")

            # Find Inbox folder inside selected mailbox
            inbox = None

            for j in range(1, selected_store.Folders.Count + 1):

                folder = selected_store.Folders.Item(j)

                print("Folder:", folder.Name)

                if folder.Name.lower() == "inbox":
                    inbox = folder
                    break

            if inbox is None:
                raise Exception(
                    f"Inbox not found in mailbox {mailbox_name}"
                )

        else:

            # Primary mailbox
            inbox = outlook.GetDefaultFolder(6)

        print(f"Processing Inbox : {inbox.Name}")

        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        total = messages.Count

        print(f"Total emails : {total}")

        for i in range(1, total + 1):

            try:

                msg = messages.Item(i)

                if not hasattr(msg, "ReceivedTime"):
                    continue

                dt = normalize_datetime(msg.ReceivedTime)

                if not dt:
                    continue

                if from_date and dt < from_date:
                    continue

                if to_date and dt > to_date:
                    continue

                processed += 1

                print(
                    f"Email: {msg.Subject[:60]} | {dt}"
                )

                attachments = msg.Attachments

                for a in range(1, attachments.Count + 1):

                    att = attachments.Item(a)

                    ext = os.path.splitext(
                        att.FileName
                    )[1].lower()

                    if ext not in [".pdf", ".doc", ".docx"]:
                        continue

                    save_path = os.path.join(
                        folder_path,
                        att.FileName
                    )

                    # Prevent overwrite
                    counter = 1
                    base, extension = os.path.splitext(save_path)

                    while os.path.exists(save_path):

                        save_path = (
                            f"{base}_{counter}{extension}"
                        )

                        counter += 1

                    att.SaveAsFile(save_path)
                    resume_metadata[os.path.basename(save_path)] = {
                        "email_date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                        "sender": msg.SenderName,
                        "subject": msg.Subject
                    }
                    downloaded += 1

                    print(
                        f"Downloaded: {os.path.basename(save_path)}"
                    )

                if progress_callback:
                    progress_callback(
                        int((i / total) * 100)
                    )

            except Exception as e:
                print(f"Email Error: {e}")

        meta_file = os.path.join(
            folder_path,
            "resume_metadata.json"
        )

        with open(meta_file, "w", encoding="utf-8") as f:
            json.dump(
                resume_metadata,
                f,
                indent=4
            )

        return processed, downloaded

    finally:
        pythoncom.CoUninitialize()