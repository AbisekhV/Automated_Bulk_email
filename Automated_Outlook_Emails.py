import pyautogui
import time
import pandas as pd
import os
import logging

# Logging setup
logging.basicConfig(filename='email_drafts.log', level=logging.INFO)

# Step 1: Load Excel Data
file_path = r'C:\\Users\\abise\\Downloads\\email_data.xlsx'
try:
    data = pd.read_excel(file_path)

    # Ensure required columns exist
    required_columns = ['Full Name', 'Company Name', 'Email Address']
    for column in required_columns:
        if column not in data.columns:
            raise ValueError(f"Column '{column}' is missing in the Excel file.")

    # Extract first name
    data['First Name'] = data['Full Name'].apply(
        lambda x: str(x).split()[0] if isinstance(x, str) else 'Unknown'
    )
except Exception as e:
    print(f"Error loading Excel file: {e}")
    logging.error(f"Error loading Excel file: {e}")
    exit()

# Step 2: Open Outlook App
try:
    print("Opening Outlook...")
    pyautogui.hotkey('win')
    time.sleep(1)
    pyautogui.write('Outlook')
    pyautogui.press('enter')
    time.sleep(10)  # Wait for Outlook to open
except Exception as e:
    print(f"Error opening Outlook: {e}")
    logging.error(f"Error opening Outlook: {e}")
    exit()

# Step 3: Create Drafts
success_count = 0
failure_count = 0

for index, row in data.iterrows():
    try:
        first_name = row['First Name']
        company_name = row['Company Name']
        email = row['Email Address']

        print(f"Creating draft for {email}...")

        # Open new email
        pyautogui.hotkey('ctrl', 'n')  # Open new email window
        time.sleep(2)

        # Fill in email details
        pyautogui.write(email)  # To address
        pyautogui.press('tab')  # Move to CC field
        #pyautogui.write('tab')  # CC address
        #pyautogui.press('tab')
        #pyautogui.press('tab')  # Move to Subject field
        pyautogui.write(f"Wishing you a growth-filled 2025")  # Subject
        pyautogui.press('tab')  # Move to Body

        # Write email body
        email_body = f"""Hi {first_name},

Wishing you and your team at {company_name} a very Happy New Year! May 2025 bring you success, happiness, and new opportunities.

Looking forward to another year of collaboration and growth together!"""
        pyautogui.write(email_body)  # Email body

        # Save as draft
        pyautogui.hotkey('ctrl', 's')  # Save the draft
        time.sleep(2)

        success_count += 1  # Track success
        print(f"Draft for {email} created successfully.")
        logging.info(f"Draft for {email} created successfully.")

    except Exception as e:
        failure_count += 1  # Track failure
        print(f"Error creating draft for {email}: {e}")
        logging.error(f"Error creating draft for {email}: {e}")

# Final summary
print(f"Process completed. Success: {success_count}, Failures: {failure_count}")
logging.info(f"Process completed. Success: {success_count}, Failures: {failure_count}")