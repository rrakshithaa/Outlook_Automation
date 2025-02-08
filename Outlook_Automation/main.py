from utils.email_processor import OutlookEmailProcessor
from config import hr_email, output_file

def main():
    processor = OutlookEmailProcessor(hr_email, output_file)
    processor.process_emails()

if __name__ == "__main__":
    main()
