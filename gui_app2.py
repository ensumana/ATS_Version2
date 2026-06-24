import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import threading
import os

from tkcalendar import DateEntry

from outlook_engine import (
    download_resumes,
    get_mailboxes
)

from ats_parser2 import process_resumes

import warnings
warnings.filterwarnings("ignore")
# ----------------------------------------------------
# UI CONFIG
# ----------------------------------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ATSApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("ATS Resume System")
        self.geometry("700x550")

        self.build_ui()
        self.load_mailboxes()

    # ------------------------------------------------
    # UI
    # ------------------------------------------------
    def build_ui(self):

        ctk.CTkLabel(
            self,
            text="ATS Resume Downloader + Parser",
            font=("Arial", 20, "bold")
        ).pack(pady=15)

        # --------------------------------------------
        # Folder Selection
        # --------------------------------------------
        self.folder_path = ctk.StringVar(
            value=os.path.join(os.getcwd(), "Downloads")
        )

        ctk.CTkEntry(
            self,
            textvariable=self.folder_path,
            width=500
        ).pack(pady=10)

        ctk.CTkButton(
            self,
            text="Browse Folder",
            command=self.browse_folder
        ).pack(pady=5)

        # --------------------------------------------
        # Mailbox Selection
        # --------------------------------------------
        ctk.CTkLabel(
            self,
            text="Outlook Mailbox"
        ).pack(pady=(15, 2))

        self.mailbox_var = ctk.StringVar()

        self.mailbox_dropdown = ctk.CTkComboBox(
            self,
            variable=self.mailbox_var,
            values=["Loading..."],
            width=500
        )

        self.mailbox_dropdown.pack(pady=5)

        ctk.CTkButton(
            self,
            text="Refresh Mailboxes",
            command=self.load_mailboxes
        ).pack(pady=5)

        # --------------------------------------------
        # Dates
        # --------------------------------------------
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=15)

        ctk.CTkLabel(
            self.frame,
            text="From Date"
        ).grid(row=0, column=0, padx=15)

        self.from_date = DateEntry(
            self.frame,
            date_pattern="dd-mm-yyyy"
        )

        self.from_date.grid(
            row=1,
            column=0,
            padx=15
        )

        ctk.CTkLabel(
            self.frame,
            text="To Date"
        ).grid(row=0, column=1, padx=15)

        self.to_date = DateEntry(
            self.frame,
            date_pattern="dd-mm-yyyy"
        )

        self.to_date.grid(
            row=1,
            column=1,
            padx=15
        )

        # --------------------------------------------
        # Progress
        # --------------------------------------------
        self.progress = ctk.CTkProgressBar(
            self,
            width=500
        )

        self.progress.set(0)

        self.progress.pack(pady=20)

        self.status = ctk.CTkLabel(
            self,
            text="Ready"
        )

        self.status.pack()

        # --------------------------------------------
        # Run Button
        # --------------------------------------------
        ctk.CTkButton(
            self,
            text="Run ATS Pipeline",
            height=45,
            width=250,
            command=self.start_pipeline
        ).pack(pady=25)

    # ------------------------------------------------
    # Load Outlook Mailboxes
    # ------------------------------------------------
    def load_mailboxes(self):

        try:

            mailboxes = get_mailboxes()

            if not mailboxes:
                mailboxes = ["No Mailboxes Found"]

            self.mailbox_dropdown.configure(
                values=mailboxes
            )

            self.mailbox_var.set(
                mailboxes[0]
            )

        except Exception as e:

            messagebox.showerror(
                "Mailbox Error",
                str(e)
            )

    # ------------------------------------------------
    # Folder Browse
    # ------------------------------------------------
    def browse_folder(self):

        folder = filedialog.askdirectory()

        if folder:
            self.folder_path.set(folder)

    # ------------------------------------------------
    # Progress
    # ------------------------------------------------
    def update_progress(self, value):

        self.progress.set(value / 100)

        self.status.configure(
            text=f"Processing: {value}%"
        )

        self.update_idletasks()

    # ------------------------------------------------
    # Start Thread
    # ------------------------------------------------
    def start_pipeline(self):

        thread = threading.Thread(
            target=self.run_pipeline
        )

        thread.daemon = True
        thread.start()

    # ------------------------------------------------
    # Main Pipeline
    # ------------------------------------------------
    def run_pipeline(self):

        try:

            folder = self.folder_path.get()

            if not os.path.exists(folder):

                os.makedirs(folder)

            from_date = datetime.strptime(
                self.from_date.get(),
                "%d-%m-%Y"
            ).replace(
                hour=0,
                minute=0,
                second=0
            )

            to_date = datetime.strptime(
                self.to_date.get(),
                "%d-%m-%Y"
            ).replace(
                hour=23,
                minute=59,
                second=59
            )

            if from_date > to_date:

                messagebox.showerror(
                    "Error",
                    "Invalid Date Range"
                )

                return

            selected_mailbox = (
                self.mailbox_var.get()
            )

            self.status.configure(
                text="Downloading resumes..."
            )

            processed, downloaded = download_resumes(
                folder,
                from_date,
                to_date,
                self.update_progress,
                selected_mailbox
            )

            self.status.configure(
                text="Parsing resumes..."
            )

            excel_file = process_resumes(
                folder
            )

            self.progress.set(1.0)

            messagebox.showinfo(
                "Completed",
                f"Mailbox: {selected_mailbox}\n\n"
                f"Processed Emails: {processed}\n"
                f"Downloaded Resumes: {downloaded}\n\n"
                f"Excel File:\n{excel_file}"
            )

            self.status.configure(
                text="Completed"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

            self.status.configure(
                text="Failed"
            )


# ----------------------------------------------------
# RUN
# ----------------------------------------------------
if __name__ == "__main__":

    app = ATSApp()
    app.mainloop()