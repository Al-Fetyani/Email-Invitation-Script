# Email Invitation Script

This script automates the process of sending email invitations using a CSV file and a Word document template.
## Installation
```
git clone https://github.com/Al-Fetyani/template_email_sender.git
```
## Prerequisites

Make sure you have Python installed on your machine. Additionally, install the required packages using Pipenv.

```
pipenv install
```
## Usage

* Populate the contacts.csv file with the necessary information.
* Customize the template.docx file for the invitation content.
* Run the script with the following command:
```
pipenv run python main.py --email your_email@gmail.com --password your_email_password
```
Note: If you don't provide the email password as a command-line argument, the script will securely prompt you to enter it.

## [Gmail Password](https://support.google.com/mail/answer/185833?hl=en)
* Go to your Google Account.
* Select Security.
* Under "Signing in to Google," select 2-Step Verification.
* At the bottom of the page, select App passwords.
* Enter a name that helps you remember where you'll use the app password.
* Select Generate.

## CSV Format
Ensure that the CSV file (contacts.csv) has the following columns (you can edit them with placeholders in template.docx file):

* salutation
* first_name
* last_name
* last_contacted
* company
* email

## Template Format
The Word document template (template.docx) should contain placeholders in square brackets that match the column names in the CSV.

Example placeholders in the template:

* [Salutation]
* [First Name]
* [Last Name]
* [Last Contacted]
* [Company Name]

License
This project is licensed under the MIT License - see the LICENSE.md file for details.

This README file provides users with information about how to use your script, what prerequisites are needed, and any other relevant details.
