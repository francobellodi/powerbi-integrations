#!/usr/bin/env python
# coding: utf-8

# In[1]:


import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from time import gmtime, strftime
import datetime
import os
import win32com.client as win32


# In[2]:


#defining initial and base report URL
initial_url = "https://app.powerbi.com/groups/9373eebd-72e5-49dd-aca6-889cab881e8d/list"
base_url = "https://app.powerbi.com/groups/9373eebd-72e5-49dd-aca6-889cab881e8d/reports/"

#Setting credentials
email = "youremail@microsoft.com"
password = "password"


# In[3]:


#Creating Driver and setting driver options
options = webdriver.ChromeOptions()

#Define directory for downloading files
prefs = {'download.default_directory' : r'C:\Users\user\projects'}
options.add_argument("--disable-extensions")
options.add_argument("--start-maximized")
options.add_argument('window-size=2560,1440')
options.add_argument("--headless")
options.add_experimental_option('prefs', prefs)

#Create Chrome driver
driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.implicitly_wait(10)

#Navigating to initial URL
driver.get(initial_url)


# In[4]:


#Creating dictionary containing all reports that need to be extracted, with their respective information and recipients
exportToPdfBase = {
    "Revenue Daily": {
        "report":"90a1e396-5790-4d18-b9ec-55baf7c77e37",
        "report_page":"ReportSection72917cadc43c76a9c8ce",
        "email_recipients": ["youremail@microsoft.com", "youremail@microsoft.com"]
    },
    "Revenue Monthly": {
        "report":"90a1e396-5790-4d18-b9ec-55baf7c77e37",
        "report_page":"ReportSection952452ec557092634076",
        "email_recipients": ["youremail@microsoft.com", "youremail@microsoft.com"]
    },
    "Receivables": {
        "report":"90a1e396-5790-4d18-b9ec-55baf7c77e37",
        "report_page":"b46561f030ed045803ec",
        "email_recipients": ["youremail@microsoft.com", "youremail@microsoft.com"]
    },
    "Revenue By Category": {
        "report":"90a1e396-5790-4d18-b9ec-55baf7c77e37",
        "report_page":"18b49249f0f077c24c0f",
        "email_recipients": ["youremail@microsoft.com", "youremail@microsoft.com"]
    },
    "Specialties": {
        "report":"90a1e396-5790-4d18-b9ec-55baf7c77e37",
        "report_page":"5e62ded0492ca000700b",
        "email_recipients": ["youremail@microsoft.com", "youremail@microsoft.com"]
    },
    "Client Base": {
        "report":"90a1e396-5790-4d18-b9ec-55baf7c77e37",
        "report_page":"49bba905c7c37db940a4",
        "email_recipients": ["youremail@microsoft.com", "youremail@microsoft.com"]
    }
}


# In[5]:


def logInMicrosoftPowerBI(email, password, driver):
    # --Pass O365 credentials finding the username/email field
    driver.find_element(By.ID, "email").send_keys(email)

    #Click the submit button
    driver.find_element(By.ID, "submitBtn").click()
    time.sleep(10)
    driver.get_screenshot_as_file("01 after email input.png")
    
    # --Pass O365 credentials finding the password field
    driver.find_element(By.ID, "i0118").send_keys(password)

    #Click the submit button
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(10)
    driver.get_screenshot_as_file("02 after password input.png")
    
    # Select not remind user
    driver.find_element(By.ID, "idBtn_Back").click()
    time.sleep(10)
    driver.get_screenshot_as_file("03 after not reminding user.png")


# In[6]:


def getReportPDF(base_url, report, report_page, driver, name):
    #Setting report page URL and navigating driver
    report_page_url = base_url + report + "/" + report_page
    driver.get(report_page_url)
    time.sleep(15)
    driver.get_screenshot_as_file(f"{name} 04 after accessing report page.png")
    
    #Finding export button and clicking
    driver.find_element(By.ID, "exportMenuBtn").click()
    time.sleep(5)
    driver.get_screenshot_as_file(f"{name} 05 after clicking export button.png")
    
    #Finding Export to PDF button and clicking
    driver.find_element('css selector', 'button[data-testid="export-to-pdf-btn"]').click()
    time.sleep(5)
    driver.get_screenshot_as_file(f"{name} 06 after clicking PDF.png")
    
    #Finding export current page only checkbox and clicking
    # Find the checkbox element and click
    checkbox = driver.find_element(By.XPATH, '//*[@name="export-only-current-page"]//input[@data-testid="pbi-checkbox-input"]')
    driver.execute_script("arguments[0].click();", checkbox)
    time.sleep(5)
    driver.get_screenshot_as_file(f"{name} 08 after clicking checkbox.png")
    
    #Finding export button and clicking
    driver.find_element(By.ID, "okButton").click()
    time.sleep(2)
    driver.get_screenshot_as_file(f"{name} 09 after clicking export button.png")
    time.sleep(60)
    
    # Renaming file for versioning control
    old_file_path = r'C:\Users\user\projects\base.pdf'
    today_date = datetime.datetime.today().strftime('%Y-%m-%d')
    new_file_name = f"{today_date}_{name}.pdf"
    new_file_path = os.path.join(r'C:\Users\user\projects', new_file_name)
    
    # Rename the file
    os.rename(old_file_path, new_file_path)
    print(f"File renamed to: {new_file_path}")
    
    return new_file_path


# In[7]:


def sendReportViaOutlook(file_path, recipients_list):
    
    # Outlook connection
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    recipients = '; '.join(recipients_list)
    
    # Email configuration
    mail.To = recipients  # Add the recipient's email
    mail.Subject = 'Your Power BI Report'
    mail.Body = 'Please find the attached report.'
    
    # Attach the file
    attachment_path = file_path
    mail.Attachments.Add(attachment_path)
    
    # Send the email
    mail.Send()
    
    print("Email sent successfully.")


# In[8]:


logInMicrosoftPowerBI(email, password, driver)


# In[10]:


for i in exportToPdfBase:
    file_path = getReportPDF(base_url, exportToPdfBase[i]["report"], exportToPdfBase[i]["report_page"], driver, i)
    sendReportViaOutlook(file_path, exportToPdfBase[i]["email_recipients"])

