# Outlook Mail Automation Project

## Project Description
This project automates the process of handling specific emails received from the HR department in Microsoft Outlook. The attached documents (in formats such as PDF, DOCX, and Excel) are parsed to extract essential details and update them into an Excel sheet.

---

## Features
- Automatically retrieve emails from Outlook based on specified criteria.
- Extract data from document attachments (PDF, DOCX, Excel).
- Parse and validate essential details such as:
  - Name
  - Number of Years of Experience
  - Phone Number
  - Email ID
- Update extracted information into an Excel sheet.
- Maintain logs for tracking automation activities and errors.

---

## Prerequisites
Ensure you have the following installed before running the project:
- **Python 3.13.0 or later**
- Microsoft Outlook application

### Required Python Packages
- `pandas`
- `openpyxl`
- `pywin32`
- `pdfplumber` (for reading PDF files)
- `python-docx` (for handling DOCX files)

Install all dependencies using the following command:
```bash
pip install pandas openpyxl pywin32 pdfplumber python-docx
```

---

## Project Setup
1. Clone this repository to your local machine:
   ```bash
   git clone <repository_url>
   ```
2. Navigate to the project directory:
   ```bash
   cd outlook-mail-automation
   ```
3. Run the main script to start the automation:
   ```bash
   python main.py
   ```

---

## Usage Instructions
1. Ensure Outlook is running and properly configured with your account.
2. Update the search criteria in the script if needed (e.g., HR email address, subject keywords).
3. Run the automation script.
4. Check the generated Excel sheet for updated details.

---

## Folder Structure
```
.
|-- main.py                  # Entry point of the project
|-- requirements.txt          # List of dependencies
|-- data                      # Folder to store Excel sheets
|   |-- extracted_details.xlsx
|-- logs                      # Folder for logs
|-- README.md                 # Project documentation
```

---

## Logs
- Logs will be generated in the `logs` folder.
- Check `automation.log` for detailed execution information.

---

## Contribution Guidelines
1. Fork the repository.
2. Create a new branch for your feature/fix.
3. Commit your changes with clear messages.
4. Submit a pull request for review.

---

## License
This project is licensed under the MIT License.

---

## Author
**Rakshitha R**

Feel free to contribute or report any issues you encounter during the usage of this project.
