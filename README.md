# Email Draft Automation Using PyAutoGUI

This script automates the creation of email drafts in Microsoft Outlook using Python and PyAutoGUI. It takes recipient details from an Excel file, customizes each email, and saves them as drafts in Outlook.

## Features
- Loads recipient details (Full Name, Company Name, Email Address) from an Excel file.
- Extracts the first name of the recipient for personalization.
- Customizes the subject and body of the email for each recipient.
- Saves each email as a draft in Microsoft Outlook.
- Logs the success and failure of draft creation for debugging and tracking purposes.

## Requirements
- Python 3.x
- Required Python libraries:
  - `pandas`
  - `openpyxl`
  - `pyautogui`
- Microsoft Outlook installed on your system.

## Setup and Usage

### Step 1: Clone the Repository
```bash
git clone https://github.com/<your-username>/<your-repo-name>.git
cd <your-repo-name>
```

### Step 2: Install Dependencies
Install the required Python libraries using pip:
```bash
pip install pandas openpyxl pyautogui
```

### Step 3: Prepare the Excel File
1. Create an Excel file named `email_data.xlsx` in the following format:
   | Full Name      | Company Name      | Email Address       |
   |----------------|-------------------|---------------------|
   | John Doe       | Example Inc.      | john@example.com    |
   | Jane Smith     | TechCorp          | jane@techcorp.com   |
2. Save this file in the `Downloads` folder or update the `file_path` variable in the script to match the file's location.

### Step 4: Run the Script
Run the script in your Python environment:
```bash
python email_draft_automation.py
```

### Step 5: Monitor Logs
The script logs the success and failure of each draft creation in `email_drafts.log`. You can review this log file to troubleshoot any issues.

## How It Works
1. **Load Data**: Reads recipient details from the Excel file and validates the required columns.
2. **Open Outlook**: Launches Microsoft Outlook using keyboard automation.
3. **Create Drafts**: Automates the creation of email drafts for each recipient, including:
   - Filling in the "To" field with the recipient's email address.
   - Writing a personalized subject and body.
   - Saving the email as a draft.
4. **Log Results**: Tracks and logs success and failure for each email draft creation.

## Notes
- Ensure Outlook is installed and properly configured on your system.
- Avoid using your system during the script execution to prevent interference with PyAutoGUI's automation.
- The script includes a delay (`time.sleep`) to accommodate for system performance and prevent issues during automation.

## Customization
- Modify the email subject and body in the script to suit your needs.
- Adjust delays (`time.sleep`) based on your system's performance.
- Update the `file_path` variable to point to your custom Excel file location.

## Troubleshooting
- If the script fails to find the required columns in the Excel file, verify the column names and format.
- Ensure Microsoft Outlook is installed and the script has sufficient time to open it.
- Review `email_drafts.log` for error details.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributions
Contributions, issues, and feature requests are welcome! Feel free to open an issue or submit a pull request.

## Author
Abisekh V (https://github.com/AbisekhV)
