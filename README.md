This Python project automates the process of logging into Power BI, exporting reports, renaming them, and sending them via email to selected recipients using Microsoft Outlook. The automation is designed to facilitate repetitive tasks and make report distribution more efficient for the project.

Features
Automated Login to Power BI: Uses Selenium WebDriver to log in to Power BI with stored credentials.
Report Downloading: Navigates through the Power BI interface to download specific reports in PDF format.
File Renaming: Automatically renames the downloaded files based on the current date and a variable-defined name.
Email Reports: Sends the reports to a list of recipients using Microsoft Outlook with attachments.
Requirements
The project requires the following dependencies to run:

Python 3.6+
Selenium for automating the browser interaction with Power BI
pywin32 for connecting to Microsoft Outlook
A compatible WebDriver for Selenium (e.g., ChromeDriver, GeckoDriver)
Install the dependencies using the following command:

bash
Copy code
pip install selenium pywin32
Setup
Clone the repository:

bash
Copy code
git clone https://github.com/yourusername/powerbi-report-exporter.git
cd powerbi-report-exporter
Set up Selenium WebDriver:

Download and install a WebDriver (e.g., ChromeDriver).
Ensure the WebDriver executable is in your system path or specify the location in your script.
Outlook Configuration:

Ensure that you are logged into Outlook on your local machine.
The script uses win32com to interact with the locally installed Outlook application. No explicit authentication is needed as it uses the active Outlook session.
Usage
Automating Report Downloads:
The script opens Power BI using Selenium, logs in, and downloads the required reports to a specific directory.

Renaming Files:
Once the reports are downloaded, the script renames them based on the current date and a predefined variable (e.g., username).

Sending Emails with Attachments:
The script sends the reports as email attachments to a list of recipients using Microsoft Outlook. The list of recipients can be modified in the script.

Example
python
Copy code
# Example of renaming a downloaded report
today_date = datetime.today().strftime('%Y-%m-%d')
name = "Núttria Report"
new_file_name = f"{today_date}_{name}.pdf"
Project Structure
bash
Copy code
.
├── Export Reports From Power BI.ipynb   # Main Jupyter notebook
├── README.md                                      # Project documentation
├── .venv/                                         # Virtual environment (optional)
└── Scripts/                                       # Folder containing Python scripts and WebDriver

Contributing
If you would like to contribute to this project, please fork the repository and submit a pull request with detailed information about the changes you propose.
