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

## [Gmail Password](https://myaccount.google.com/apppasswords?pli=1&rapt=AEjHL4Po7vgONUOI9GvrgnMxd85STlPGEVkCCPSaoBvFNKnFV1RVYzcz0fB_y2uO_37g5p0KRt0ZoYzip4KqvR4GY-IbsRzEtqsFLv8sDgsoew6m6-kEw8Q)
* Go to Manage Google Account.
* Search for App passwords.
* Add name to remember what is password for.
* Click create.
* Copy.

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
