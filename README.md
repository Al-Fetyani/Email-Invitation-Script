# Email Invitation Script

This script automates the process of sending email invitations using a CSV file and a Word document template.

## Prerequisites

Make sure you have Python installed on your machine. Additionally, install the required packages using Pipenv.

```bash
pipenv install

# Usage

1- Populate the contacts.csv file with the necessary information.
2- Customize the template.docx file for the invitation content.
3- Run the script with the following command:
pipenv run python script_name.py --email your_email@gmail.com --password your_email_password

Note: If you don't provide the email password as a command-line argument, the script will securely prompt you to enter it.

CSV Format
Ensure that the CSV file (contacts.csv) has the following columns:

salutation
first_name
last_name
last_contacted
company
email
Template Format
The Word document template (template.docx) should contain placeholders in square brackets that match the column names in the CSV.

Example placeholders in the template:

[Salutation]
[First Name]
[Last Name]
[Last Contacted]
[Company Name]
License
This project is licensed under the MIT License - see the LICENSE.md file for details.

Acknowledgments
pandas
smtplib
docx
email
argparse


3. Replace `script_name.py` with the actual name of your Python script.

4. Customize the content and sections of the README file as needed.

5. Save and commit the `README.md` file to your GitHub repository.

This README file provides users with information about how to use your script, what prerequisites are needed, and any other relevant details.
