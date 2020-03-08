from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime

import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import os
from pathlib import Path, PureWindowsPath

import sys

class Bot():
    def __init__(self):
        self.driver=webdriver.Chrome(executable_path="chromedriver\chromedriver.exe")

    def gogogo(self):
        self.driver.get('https://corp.csloxinfo.com/')
        
        self.driver.implicitly_wait(10)

        user_name=self.driver.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td[2]/label/input')
        user_name.click()
        user_name.send_keys('172814')
        #user_name.send_keys(Keys.TAB)
        
        pwdtxt=self.driver.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/label/input')
        pwdtxt.click()
        pwdtxt.send_keys('1728141234')

        login_btn=self.driver.find_element_by_xpath('//*[@id="button"]')
        login_btn.click()

        self.driver.implicitly_wait(10)

        main_link=self.driver.find_element_by_xpath('//*[@id="myTable"]/tbody/tr[1]/td[1]/center/strong/a')
        main_link.click()

        # txt_status=self.driver.find_element_by_xpath('//*[@id="jDash"]/table/tbody/tr/td[1]/div[1]/div/table/tbody/tr[4]/td[2]/text()')
        # if str(txt_status).lower()=='up':
        img_link_1=self.driver.find_element_by_xpath('//*[@id="zoomGraphImage"]')
        img_link_1.click()

        img_link_week1=self.driver.find_element_by_xpath('/html/body/table[3]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/a/img')
        img_link_week1.click()


        report_date=str(datetime.date.today())

        #create save path
        # directory = PureWindowsPath('../weekly_reports/')
        # filename = 'weekly_report_' + report_date + '.png'
        # if not os.path.isdir(directory):
        #     os.mkdir(directory)
        
        file_path =  './weekly_reports/weekly_report_' + report_date + '.png'

        ## check close button from popup page, if can click system will capture screen to file
        try:
            element = WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/div/div/div/div[2]/div[4]/a'))
            )
            #graph image
            #//*[@id="zoomGraphImage"]

            #close button
            #/html/body/div[2]/div[3]/div/div/div/div[2]/div[4]/a

            #EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div/div/div/div[2]/div[4]/a'))
            
            savefile= file_path
            self.driver.get_screenshot_as_file(savefile)
          
        finally:
            self.driver.quit()

        #     # send email 
            subject = "Internet Link Weekly Report : " + report_date
            body = "Automatic Internet Link Weekly Report : " + report_date
            sender_email = "printer@nclthailand.com"
            receiver_email = "it.infra@nclthailand.com"
            password = 'Pass@worD!!'

            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = subject
            message["Bcc"] = receiver_email  # Recommended for mass emails

            # Add body to email
            message.attach(MIMEText(body, "plain"))

            #filename = 'C:\weekly_reports\internet_weekly_report_' + report_date + '.png' # In same directory as script

            filename1=file_path

            # Open PDF file in binary mode
            with open(filename1, "rb") as attachment:
                # Add file as application/octet-stream
                # Email client can usually download this automatically as attachment
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            # Encode file in ASCII characters to send by email    
            encoders.encode_base64(part)

            # Add header as key/value pair to attachment part
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {filename1}",
            )

            # Add attachment to message and convert message to string
            message.attach(part)
            text = message.as_string()

            # Log in to server using secure context and send email
            
           # context = ssl.create_default_context()
            with smtplib.SMTP(host="smtp.office365.com", port= 587) as server:
                server.starttls()
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)


bot = Bot()
bot.gogogo()

sys.exit()