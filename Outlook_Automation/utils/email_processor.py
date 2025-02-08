import win32com.client
from Outlook_Automation.utils.file_extractor import FileExtractor
from Outlook_Automation.utils.detail_extractor import DetailExtractor


class OutlookEmailProcessor:
    def __init__(self, sender_email, file_path):
        self.sender_email = sender_email
        self.file_path = file_path

    def process_emails(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)  # Inbox
            messages = inbox.Items

            for message in messages:
                if self._is_target_email(message):
                    print(f"Processing email: {message.Subject}")
                    for attachment in message.Attachments:
                        self._process_attachment(attachment)
        except Exception as e:
            print(f"Error connecting to Outlook: {e}")

    def _is_target_email(self, message):
        try:
            sender = message.Sender
            actual_email = sender.Address if not sender.GetExchangeUser() else sender.GetExchangeUser().PrimarySmtpAddress
            return actual_email.lower() == self.sender_email.lower()
        except:
            return False

    def _process_attachment(self, attachment):
        try:
            file_content = attachment.Content
            file_type = attachment.FileName.split('.')[-1].lower()
            details = FileExtractor.extract_details(file_content, file_type)
            if details:
                DetailExtractor.update_file(self.file_path, details)
        except Exception as e:
            print(f"Error processing attachment: {e}")
