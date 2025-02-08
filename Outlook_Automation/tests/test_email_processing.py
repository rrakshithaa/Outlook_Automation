import unittest
from unittest.mock import MagicMock, patch
from utils.email_processor import OutlookEmailProcessor

class TestOutlookEmailProcessor(unittest.TestCase):

    @patch("win32com.client.Dispatch")
    def setUp(self, mock_dispatch):
        self.processor = OutlookEmailProcessor("test@example.com", "test_file.xlsx")
        self.mock_outlook = mock_dispatch.return_value.GetNamespace.return_value
        self.mock_inbox = self.mock_outlook.GetDefaultFolder.return_value
        self.mock_message = MagicMock()
        self.mock_inbox.Items = [self.mock_message]

    def test_process_emails_valid_sender(self):
        """Test processing emails when sender matches."""
        self.mock_message.Sender.Address = "test@example.com"
        self.mock_message.Attachments = []
        self.mock_message.Body = "Candidate Name: John Doe\nExperience: 5 years\nPhone Number: +1234567890\nEmail ID: john.doe@example.com"

        with patch("utils.detail_extractor.DetailExtractor.update_file") as mock_update_file:
            self.processor.process_emails()
            mock_update_file.assert_called_once()
            print("Valid sender email processed successfully!")

    def test_process_emails_invalid_sender(self):
        """Test skipping emails when sender does not match."""
        self.mock_message.Sender.Address = "other@example.com"
        with patch("utils.detail_extractor.DetailExtractor.update_file") as mock_update_file:
            self.processor.process_emails()
            mock_update_file.assert_not_called()
            print("Skipped processing invalid sender email.")

    def test_process_attachment(self):
        """Test processing attachments with valid details."""
        self.mock_message.Sender.Address = "test@example.com"
        mock_attachment = MagicMock()
        mock_attachment.FileName = "sample.docx"
        mock_attachment.Content = b"Sample attachment content"
        self.mock_message.Attachments = [mock_attachment]

        with patch("utils.file_extractor.FileExtractor.extract_details", return_value={"Candidate Name": "John Doe"}):
            with patch("utils.detail_extractor.DetailExtractor.update_file") as mock_update_file:
                self.processor.process_emails()
                mock_update_file.assert_called_once()
                print("Attachment processed successfully!")

if __name__ == "__main__":
    unittest.main()
