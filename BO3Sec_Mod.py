# -*- coding: utf-8 -*-
"""
Created on Sat May  2 20:38:08 2020

@author: Nandal
"""
import datetime
from datetime import timedelta
from selenium import webdriver
from time import localtime, strftime
import time
import sys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from tkinter import ttk
from ttkthemes import ThemedTk
import tkinter as tk
import pyttsx3
from playsound import playsound
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import pandas as pd
from tkinter import Canvas
import zipfile
import requests

try:
  driver = webdriver.Chrome(executable_path=r'chromedriver.exe')
except:
  # get a self updating webdriver
  print('*'*45)
  print('*'*45)
  print(''*45)
  print('Please wait, updating chrome driver, just go to your browser and paste the chrome version (for e.g. - 99.0.4197.75) and hit enter\n')
  ans = 'Y'
  while ans == 'Y':
    url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_'
    url_file = 'https://chromedriver.storage.googleapis.com/'
    file_name = 'chromedriver_win32.zip'
    version = input('Now Enter Current Chrome Browser Version from :(3 dots) -> Help -> About Google Chrome: ')
    version_test = version.split('.')[:3]
    version_test = '.'.join(version_test)  
    version_response = requests.get(url + version_test)
    if 'Error' not in version_response.text:
        file = requests.get(url_file + version_response.text + '/' + file_name)
        with open(file_name, "wb") as code:
            code.write(file.content)
        print('\nSuccessfully downloaded new version, now extracting zip.....')
        with zipfile.ZipFile(file_name, 'r') as zip:
            # printing all the contents of the zip file
            zip.printdir()
          
            # extracting all the files
            print('\nExtracting all the files now...')
            zip.extractall()
            print('*'*45)
            print('*'*45)            
            print('\nDone! Re-starting program!! Please wait...')
            print('*'*45)
            print('*'*45)                
            driver = webdriver.Chrome(executable_path=r'chromedriver.exe')
        ans = 'n'
    else:
      answer = input('\nDid you enter the version correctly, without any leading or trailing spaces? Try again?: [Y/N] ?')
      if answer == '':
        ans = 'Y'
      else:
        ans = answer[0].upper()
 
global data
global data_banks
data = pd.read_excel('MASTER.xlsx')
data = data.set_index('key')
engine = pyttsx3.init()
engine.setProperty('rate', 165)

data_banks = pd.read_excel('BANKS-NOS-DOS.xlsx')
data_banks = data_banks.set_index('key')

def excel_updater():
  global data
  global data_banks
  data = pd.read_excel('MASTER.xlsx')
  data = data.set_index('key')
  data_banks = pd.read_excel('BANKS-NOS-DOS.xlsx')
  data_banks = data_banks.set_index('key')
  
def first():
  None

def active():
  driver.switch_to.window(driver.window_handles[-1])

def start():
  try:
    playsound('button.wav')
  except:
    None
    
def over():
  try:
    playsound('over.wav')
  except:
    None
  return

if int(strftime("%H", localtime())) > 4 and int(strftime("%H", localtime())) < 8:
  engine = pyttsx3.init()
  engine.say('Hey, Early bird, Good morning.')
  engine.runAndWait()
elif int(strftime("%H", localtime())) >= 22 or int(strftime("%H", localtime())) <= 4:
  engine = pyttsx3.init()
  engine.say('You should not work so late. Anyways, I am right here')
  engine.runAndWait()
elif int(strftime("%H", localtime())) > 17 and int(strftime("%H", localtime())) < 22:
  engine = pyttsx3.init()
  engine.say('Good evening')
  engine.runAndWait()
else:
  engine = pyttsx3.init()
  engine.say('What a beautiful day')
  engine.runAndWait()

def CMS_login():
  active()
  driver.delete_all_cookies()
  driver.get('https://cms.rbi.org.in/ro/app/login/login')
  driver.maximize_window()
  #waiting to enter manually username and password in popup dialog
  if driver.title[:3] == '502':
    engine.say('5. 0. 2. error. CMS, god damn it')
    engine.runAndWait()
  else:
    engine.say('Please log in')
    engine.runAndWait()

def logout():
  active()
  driver.delete_all_cookies()
  driver.get('https://cms.rbi.org.in/ro/app/login/logout')
  if driver.title[:3] == '502':
    engine.say('5. 0. 2. error. CMS, god damn it')
    engine.runAndWait()
  else:
    engine.say('You are now logged out. Cheers!')
    engine.runAndWait()

def Search():
  start()
  case_nos = entryS.get().split()
  active()
  for i in range(len(case_nos)):
    driver.execute_script('''window.open("about:blank",'_newtab_home');''')
    driver.switch_to.window(driver.window_handles[-1])
    driver.get('https://cms.rbi.org.in/ro/app/CRMNextObject/SearchAction/Case')
    in_field = driver.find_element_by_xpath('//*[@id="objectWrapper"]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div[1]/div[1]/div/div/div/input')
    in_field.send_keys(case_nos[i] + Keys.ENTER)
  over()
    
def SearchE():
  start()
  case_nos = entryS.get().split()
  active()
  for i in range(len(case_nos)):
    driver.execute_script('''window.open("about:blank",'_newtab_home');''')
    driver.switch_to.window(driver.window_handles[-1])
    driver.get('https://cms.rbi.org.in/ro/app/CRMNextObject/SearchAction/Case')
    in_field = driver.find_element_by_xpath('//*[@id="objectWrapper"]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div[1]/div[1]/div/div/div/input')
    in_field.send_keys(case_nos[i] + Keys.ENTER)
    try:
      WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Edit"]'))).click()
    except:
      engine.say('I have waited for more than 40 seconds, please check your internet connection !')
      engine.runAndWait()
  over()
  
def refresh_all():
  start()
  active()
  
  engine.say('Refreshing {} tabs.'.format(len(driver.window_handles)))
  engine.runAndWait()
  
  for i in range(len(driver.window_handles)):
    driver.refresh()
    driver.switch_to.window(driver.window_handles[i])
  over()

def open_drafts_SFA():
  start()
  active()
  # number of cases on any page
  records = driver.find_elements_by_xpath('//div[@class="showrecords"]')[0].text
  records = records.split()
  hyphen_index = records[1].index('-')
  records = records[1]
  first = int(records[:hyphen_index])
  second = int(records[hyphen_index+1:])
  cases_count = second-first+1

  xpath_statuses = ['//div[@data-autoid="StatusCode_' + str(i) + '"]' for i in range(cases_count)]
  
  status_list = [driver.find_elements_by_xpath(xpath_statuses[x])[0].text for x in range(cases_count)]
  
  indices_cases = [i for i,x in enumerate(status_list) if x =='Sent for approval']
  
  if len(indices_cases) != 0:
    print('Found {} drafts awaiting approval on page.... Opening!'.format(len(indices_cases)))
    engine.say('Found {} drafts awaiting approval on page.... Opening!'.format(len(indices_cases)))
    engine.runAndWait()
    
    xpath_links_all = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[2]/div/div[2]/div/a' for x in indices_cases]
   
    xpath_edit_menu = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[1]' for x in indices_cases] 
   
    # case_links = []
    
    for i in range(len(indices_cases)):
      while True:
        edit_menu= WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, xpath_edit_menu[i])))
        ActionChains(driver).move_to_element(edit_menu).perform()
        try:
          temp = driver.find_element_by_xpath(xpath_links_all[i])
          ActionChains(driver).move_to_element(temp).key_down(Keys.CONTROL).click().key_up(Keys.CONTROL).perform()
          # case_links.append(temp)
          break
        except NoSuchElementException:
          continue
  
  #   for i in range(len(indices_cases)):
  #     driver.switch_to.window(driver.window_handles[-1])
  #     driver.execute_script("window.open('" + case_links[i] + "','_newtab_home');")
    over()
  else:
    engine.say('No such cases found!! Check again.')
    engine.runAndWait()

# statuses = ['Complaint Closed','Sent to DO','BO Decision','Sent back to DO','Sent to','Secretary BO','Sent to other office','Assign to other Regulatory bodies','Issue Advisory','Sent to Other Regulatory Bodies','Assign to other RBI Department','Sent to Closing Authority','Appeal Closed','Sent Back to CEPC DO','New Complaint','New Appeal','Re Opened','Not a Complaint']

def open_New_C():
  start()
  active()
  # number of cases on any page
  records = driver.find_elements_by_xpath('//div[@class="showrecords"]')[0].text
  records = records.split()
  hyphen_index = records[1].index('-')
  records = records[1]
  first = int(records[:hyphen_index])
  second = int(records[hyphen_index+1:])
  cases_count = second-first+1
  
  xpath_links_all = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[1]/div[1]/div/div/span/a' for x in range(cases_count)]
 
  case_links = [driver.find_element_by_xpath(xpath_links_all[x]) for x in range(cases_count)]
  
  xpath_statuses = ['//div[@data-autoid="StatusCode_' + str(i) + '"]' for i in range(cases_count)]
  
  status_list = [driver.find_elements_by_xpath(xpath_statuses[x])[0].text for x in range(cases_count)]
  
  #opening all New Complaints in new tab
  
  indices_cases = [i for i,x in enumerate(status_list) if x =='New Complaint']
  
  if len(indices_cases) != 0:
    print('Found {} New Complaints on page.... Opening!'.format(len(indices_cases)))
    engine.say('Found {} New Complaints on page.... Opening!'.format(len(indices_cases)))
    engine.runAndWait()
    for i in indices_cases:
      ActionChains(driver).move_to_element(case_links[i]).key_down(Keys.CONTROL).click().key_up(Keys.CONTROL).perform()
      # driver.execute_script('''window.open("about:blank",'_newtab_home');''')
      # driver.switch_to.window(driver.window_handles[-1])
      # driver.get(case_links[i])
    over()
  else:
    engine.say('No such cases found!! Check again.')
    engine.runAndWait()

def open_BO_D():
  start()
  active()
  # number of cases on any page
  records = driver.find_elements_by_xpath('//div[@class="showrecords"]')[0].text
  records = records.split()
  hyphen_index = records[1].index('-')
  records = records[1]
  first = int(records[:hyphen_index])
  second = int(records[hyphen_index+1:])
  cases_count = second-first+1
  
  xpath_statuses = ['//div[@data-autoid="StatusCode_' + str(i) + '"]' for i in range(cases_count)]
  
  status_list = [driver.find_element_by_xpath(xpath_statuses[x]).text for x in range(cases_count)]
  
  indices_cases = [i for i,x in enumerate(status_list) if x =='BO Decision']
  
  if len(indices_cases) != 0:
    print('Found {} B O Decisions on page.... Opening!'.format(len(indices_cases)))
    engine.say('Found {} B O Decisions on page.... Opening!'.format(len(indices_cases)))
    engine.runAndWait()
    
    xpath_links_all = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[2]/div/div[2]/div/a' for x in indices_cases]
   
    xpath_edit_menu = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[1]' for x in indices_cases] 
   
    # case_links = []
    
    for i in range(len(indices_cases)):
      while True:
        edit_menu= WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, xpath_edit_menu[i])))
        ActionChains(driver).move_to_element(edit_menu).perform()
        try:
          temp = driver.find_element_by_xpath(xpath_links_all[i])
          ActionChains(driver).move_to_element(temp).key_down(Keys.CONTROL).click().key_up(Keys.CONTROL).perform()
          # temp = driver.find_element_by_xpath(xpath_links_all[i]).get_attribute('href')
          # case_links.append(temp)
          break
        except NoSuchElementException:
          continue
  
  #   for i in range(len(indices_cases)):
  #     driver.switch_to.window(driver.window_handles[-1])
  #     driver.execute_script("window.open('" + case_links[i] + "','_newtab_home');")
    over()
  else:
    engine.say('No such cases found!! Check again.')
    engine.runAndWait()

def open_all():
  start()
  active()
  # number of cases on any page
  records = driver.find_elements_by_xpath('//div[@class="showrecords"]')[0].text
  records = records.split()
  hyphen_index = records[1].index('-')
  records = records[1]
  first = int(records[:hyphen_index])
  second = int(records[hyphen_index+1:])
  cases_count = second-first+1
  
  xpath_links_all = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[1]/div[1]/div/div/span/a' for x in range(cases_count)]
 
  # case_links = [driver.find_elements_by_xpath(xpath_links_all[x])[0].get_attribute('href') for x in range(cases_count)]
  case_links = [driver.find_element_by_xpath(xpath_links_all[x]) for x in range(cases_count)]
  
  if len(case_links) != 0:
    print('Found {} Cases on page.... Opening!'.format(len(case_links)))
    engine.say('Found {} Cases on page.... Opening!'.format(len(case_links)))
    engine.runAndWait()
    for i in range(len(case_links)):
      ActionChains(driver).move_to_element(case_links[i]).key_down(Keys.CONTROL).click().key_up(Keys.CONTROL).perform()
      # driver.execute_script('''window.open("about:blank",'_newtab_home');''')
      # driver.switch_to.window(driver.window_handles[-1])
      # driver.get(case_links[i])
    over()
  else:
    engine.say('No such cases found!! Check again.')
    engine.runAndWait()

def edit_all():
  start()
  active()
  # number of cases on any page
  records = driver.find_elements_by_xpath('//div[@class="showrecords"]')[0].text
  records = records.split()
  hyphen_index = records[1].index('-')
  records = records[1]
  first = int(records[:hyphen_index])
  second = int(records[hyphen_index+1:])
  cases_count = second-first+1
  
  if cases_count != 0:
    print('Found {} cases on page.... Opening!'.format(cases_count))
    engine.say('Found {} cases on page.... Opening!'.format(cases_count))
    engine.runAndWait()
    
    xpath_links_all = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[2]/div/div[2]/div/a' for x in range(cases_count)]
   
    xpath_edit_menu = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[1]' for x in range(cases_count)] 
   
    # case_links = []
    
    for i in range(cases_count):
      while True:
        edit_menu= WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, xpath_edit_menu[i])))
        ActionChains(driver).move_to_element(edit_menu).perform()
        try:
          temp = driver.find_element_by_xpath(xpath_links_all[i])
          # temp = driver.find_element_by_xpath(xpath_links_all[i]).get_attribute('href')
          ActionChains(driver).move_to_element(temp).key_down(Keys.CONTROL).click().key_up(Keys.CONTROL).perform()
          # case_links.append(temp)
          break
        except NoSuchElementException:
          continue
  
  #   for i in range(cases_count):
  #     driver.switch_to.window(driver.window_handles[-1])
  #     driver.execute_script("window.open('" + case_links[i] + "','_newtab_home');")
    over()
  else:
    engine.say('No cases found!! Check again.')
    engine.runAndWait()


# put up any maintainable complaint with format

# 11.3.a - all

def resolved():
  start()
  active()
  ct = NMC_clause(Sec113a)
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
    #check if already clicked on edit and click on edit if not and pick complaint data
      try:
        driver.find_element_by_xpath('//*[@id="object-action-button"]/a[1]').click()
      except:
        None
      # #click on assessment
      # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="1"]/span'))).click()
      # #click on Sent to BO
      # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="state_0"]/span'))).click()
    
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #click on sent to BO
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent to BO"]'))).click()
    
      while True:
        try:
          #select proposal box - M/NM and send clause value to be typed
          prop_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select')))
          driver.execute_script("arguments[0].click();", prop_box)
          ActionChains(driver).move_to_element(prop_box).click(prop_box).perform()             
          time.sleep(0.1)
          prop_box.send_keys('m')          
          #delete previous value if any in proposed clause
          prop_cls = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
          ActionChains(driver).move_to_element(prop_cls).click(prop_cls).perform()             
          time.sleep(0.1)
          prop_cls.clear()
          prop_cls.send_keys(ct[0])   
          #wait for clause to come and select
          WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()
          break
        except (NoSuchElementException, TimeoutException):
          continue
        
      while True:
        try:
          #select FRC No/Yes
          if ct[0] == '9 (3) (a)':
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select'))).send_keys('y' + Keys.TAB + Keys.TAB)
          else:
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select'))).send_keys('n' + Keys.TAB + Keys.TAB)
          #delete previous value if any in complaint status
          status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/input')))
          ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
          time.sleep(0.1)
          status_box.clear()          
          status_box.send_keys('Complaint in Process')
          #wait for status to come and select
          WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
          
          ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(ct[1]).perform()
          break
        except (NoSuchElementException, TimeoutException):
          continue
    except IndexError:
        over()
  
def NMC_clause(fun):
  val=0
  val=fun()
  return val
# NMC_clause FUNCTION THAT GIVESS CLAUSE VALUE TO ANOTHER FUNCTION NMC_clause
def Sec113a():
  ct=[]
  ct.append('11(3A)')
  ct.append(data.loc['113A'].to_dict()['value'])
  return ct
def _91():
  ct=[]
  ct.append('9(1)')
  ct.append(data.loc[91].to_dict()['value'])
  return ct
def _92():
  ct=[]
  ct.append('9 (2)')
  ct.append(data.loc[92].to_dict()['value'])
  return ct
def _93a():
  ct=[]
  ct.append('9 (3) (a)')
  ct.append(data.loc['93a'].to_dict()['value'])
  return ct
def _93b():
  ct=[]
  ct.append('9 (3) (b)')
  ct.append(data.loc['93b'].to_dict()['value'])
  return ct
def _93c():
  ct=[]
  ct.append('9 (3) (c)')
  ct.append(data.loc['93c'].to_dict()['value'])
  return ct
def _7191b():
  ct=[]
  ct.append('7(1) 9(1)')
  ct.append(data.loc['7191b'].to_dict()['value']) 
  return ct
def _7191c():
  ct=[]
  ct.append('7(1) 9(1)')
  ct.append(data.loc['7191c'].to_dict()['value'])
  return ct
def _8182():
  ct=[]
  ct.append('Grounds of complaint is/are not covered under Clause 8 of the Ombudsman Scheme, 2006')
  ct.append(data.loc[8182].to_dict()['value']) 
  return ct

# NMC FUNCTION THAT ACCEPTS CLAUSE VALUE FROM ANOTHER FUNCTION NMC_clause_.. AND PUTS UP TO BO
def NMC(fun):
  start()
  active()
  ct = NMC_clause(fun)
  #check if already clicked on edit and click on edit if not and pick complaint data
  try:
    driver.find_element_by_xpath('//*[@id="object-action-button"]/a[1]').click()
  except:
    None
  # #click on assessment
  # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="1"]/span'))).click()
  # #click on Sent to BO
  # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="state_0"]/span'))).click()

  WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
  #click on sent to BO
  WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent to BO"]'))).click()

  while True:
    try:
      #select proposal box - M/NM and send clause value to be typed
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n' + Keys.TAB)
      #delete previous value if any in proposed clause
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input'))).send_keys(Keys.CONTROL + 'a' + Keys.BACKSPACE)
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input'))).send_keys(ct[0])
      #wait for clause to come and select
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()
      break
    except (NoSuchElementException, TimeoutException):
      continue
    
  while True:
    try:
      #select FRC No/Yes
      if ct[0] == '9 (3) (a)':
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select'))).send_keys('y' + Keys.TAB + Keys.TAB)
      else:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select'))).send_keys('n' + Keys.TAB + Keys.TAB)
      #delete previous value if any in complaint status
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/input'))).send_keys(Keys.CONTROL + 'a' + Keys.BACKSPACE)
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/input'))).send_keys('Complaint with RBI')
      #wait for status to come and select
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(ct[1]).perform()
      break
    except (NoSuchElementException, TimeoutException):
      continue
  
  # comment_field = driver.find_element_by_xpath('//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div/div/div/div/div[2]/div/div/textarea')
  # comment_field.send_keys(ct[1])
  over()

#CLOSE ALL SENT FOR APPROVAL AS NAC
def CASFA():
  start()
  active()
  clause_text = data.loc['draft-NAC'].to_dict()['value']
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on NAC
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Not a Complaint"]'))).click()
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

# 13 1.a complaints sent to secretary BO

def resolved_SP():
  start()
  active()
  ct = NMC_clause(Sec113a)
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
    #check if already clicked on edit and click on edit if not and pick complaint data
      try:
        driver.find_element_by_xpath('//*[@id="object-action-button"]/a[1]').click()
      except:
        None
      # #click on assessment
      # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="1"]/span'))).click()
      # #click on Sent to BO
      # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="state_0"]/span'))).click()
    
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #click on sent to BO
      WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent to BO"]'))).click()
    
      while True:
        try:
          #select proposal box - M/NM and send clause value to be typed
          prop_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select')))
          driver.execute_script("arguments[0].click();", prop_box)
          ActionChains(driver).move_to_element(prop_box).click(prop_box).perform()             
          time.sleep(0.1)
          prop_box.send_keys('m')          
          #delete previous value if any in proposed clause
          prop_cls = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
          ActionChains(driver).move_to_element(prop_cls).click(prop_cls).perform()             
          time.sleep(0.1)
          prop_cls.clear()
          prop_cls.send_keys(ct[0])   
          #wait for clause to come and select
          WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()
          break
        except (NoSuchElementException, TimeoutException):
          continue
        
      while True:
        try:
          #select FRC No/Yes
          if ct[0] == '9 (3) (a)':
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select'))).send_keys('y' + Keys.TAB + Keys.TAB)
          else:
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select'))).send_keys('n' + Keys.TAB + Keys.TAB)
          #delete previous value if any in complaint status
          status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/input')))
          ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
          time.sleep(0.1)
          status_box.clear()
          status_box.send_keys('Complaint in Process')
          #wait for status to come and select
          WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
          
          ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(ct[1]).perform()
          driver.find_element_by_xpath('*//span[text()="Save and Proceed"]').click()
          break
        except (NoSuchElementException, TimeoutException):
          continue
    except IndexError:
        over()

#### SCRIPT FOR PROCESSING MAINTAINABLE CASES - SCRUTINY
def maintainable():
  start()
  active()
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()
  
  bank_in_cms = driver.find_element_by_xpath('*//span[@data-autoid="cust_5958_ctrl"]').text
  no_name = data_banks.loc[bank_in_cms].to_dict()['value-NO'] 
  # do_name = data_banks.loc[bank_in_cms].to_dict()['value-DO']
  status = 'Complaint with Bank'
  
  if bank_in_cms == 'STATE BANK OF INDIA':
    engine.say('There are 2 Nodal Officers for {}, please select proper button.'.format(bank_in_cms))
    engine.runAndWait()
    over()
    raise ValueError
  else:
    timeout = time.time() + 30
      #clicking on first edit button
    while True:
      try:
        # finding edit menu hover icon and going to it
        edit_menu= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="action-hover"]')))
        ActionChains(driver).move_to_element(edit_menu).perform()
        #clicking on edit-icon
        edit_icon= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Edit"]')))
        driver.get(edit_icon.get_attribute('href'))
        break
      except (NoSuchElementException, TimeoutException):
        if  time.time() > timeout:
          engine.say('Re fresh the page and try again. Seems like a network problem.')
          engine.runAndWait()
          over()
          raise ValueError
        else:
          continue
    active()
    #scrolling down 
    driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
    time.sleep(2.5)
    
    timeout = time.time() + 120
    while True:
      try:  
        no_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="ISSUEREQ_OWNER"]')))
        ActionChains(driver).move_to_element(no_name_box).click(no_name_box).perform()
        time.sleep(0.5)
        no_name_box.send_keys(Keys.CONTROL,'a')
        time.sleep(0.1)
        ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.1)
        ActionChains(driver).send_keys(no_name).perform()
        #selecting NO name
        no_elem = '*//td[text()=' + '"' + no_name + '"' + ']'
        WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, no_elem))).click()
        break
      except (NoSuchElementException, TimeoutException):
        if  time.time() > timeout:
          engine.say('Please go back, Re fresh the page and try again. Seems like a network problem.')
          engine.runAndWait()
          over()
          raise ValueError
        else:
          continue
    #click on Save
    WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save"]'))).click()
    # time.sleep(0.4)
    # #click on Close to go back to case
    # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Close"]'))).click()
    # time.sleep(0.4)
    # #click on Edit to assign it to the DO
    # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//i[@class="icon icon-edit"]'))).click()
    # time.sleep(0.4)
    # #click on assessment
    # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
    # #click on sent back to DO or sent to DO
    # try:
    #   WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent back to DO"]'))).click()
    # except:
    #   WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent to DO"]'))).click()
      
    # timeout = time.time() + 120
    # while True:
    #   try:  
    #     #DO Name box
    #     do_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="cust_5788"]')))
    #     ActionChains(driver).move_to_element(do_name_box).click(do_name_box).perform()
    #     time.sleep(0.1)
    #     do_name_box.send_keys(Keys.CONTROL,'a')
    #     time.sleep(0.1)
    #     ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
    #     time.sleep(0.1)
    #     ActionChains(driver).send_keys(do_name).perform()
    #     #selecting DO name
    #     do_elem = '*//td[text()=' + '"' + do_name + '"' + ']'
    #     WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, do_elem))).click()
    #     break
    #   except (NoSuchElementException, TimeoutException):
    #     if  time.time() > timeout:
    #       engine.say('Re fresh the page and try again. Seems like a network problem.')
    #       engine.runAndWait()
    #       over()
    #       raise ValueError
    #     else:
    #       continue
    
    # timeout = time.time() + 120
    # while True:
    #   try:
    #     #delete previous value if any in complaint status
    #     prop_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="cust_6008"]')))
    #     prop_box.send_keys(Keys.CONTROL,'a')
    #     prop_box.send_keys(Keys.BACKSPACE)
    #     #send Complaint with bank value
    #     prop_box.send_keys(status)
    #     status_elem = '*//td[text()=' + '"' + status + '"' + ']'
    #     #wait for status to come and select
    #     WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, status_elem))).click()
    #     break
    #   except (NoSuchElementException, TimeoutException):
    #     if  time.time() > timeout:
    #       engine.say('Re fresh the page and try again. Seems like a network problem.')
    #       engine.runAndWait()
    #       over()
    #       raise ValueError
    #     else:
    #       continue
    #   # click on Save and Proceed
    # driver.find_element_by_xpath('*//span[text()="Save and Proceed"]').click()
    over()
      

# def IBK():
#   return_list = []
#   no = data_banks.loc['INDIAN BANK']['value-NO'][0]
#   return_list.append(no)
#   do = data_banks.loc['INDIAN BANK']['value-DO'][0]
#   return_list.append(do)
#   return return_list
# def IBN():
#   return_list = []
#   no = data_banks.loc['INDIAN BANK']['value-NO'][1]
#   return_list.append(no)
#   do = data_banks.loc['INDIAN BANK']['value-DO'][1]
#   return_list.append(do)
#   return return_list
def SBIDEL():
  return_list = []
  no = data_banks.loc['STATE BANK OF INDIA']['value-NO'][0]
  return_list.append(no)
  do = data_banks.loc['STATE BANK OF INDIA']['value-DO'][0]
  return_list.append(do)
  return return_list
def SBICHD():
  return_list = []
  no = data_banks.loc['STATE BANK OF INDIA']['value-NO'][1]
  return_list.append(no)
  do = data_banks.loc['STATE BANK OF INDIA']['value-DO'][1]
  return_list.append(do)
  return return_list


def maintainable_bank(func):
  start()
  active()
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()
  
  get_list = func()
  
  bank_in_cms = driver.find_element_by_xpath('*//span[@data-autoid="cust_5958_ctrl"]').text
  no_name = get_list[0] 
  # do_name = get_list[1]
  # status = 'Complaint with Bank'
  
  if bank_in_cms == 'STATE BANK OF INDIA':
    timeout = time.time() + 30
      #clicking on first edit button
    while True:
      try:
        # finding edit menu hover icon and going to it
        edit_menu= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="action-hover"]')))
        ActionChains(driver).move_to_element(edit_menu).perform()
        #clicking on edit-icon
        edit_icon= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Edit"]')))
        driver.get(edit_icon.get_attribute('href'))
        break
      except (NoSuchElementException, TimeoutException):
        if  time.time() > timeout:
          engine.say('Re fresh the page and try again. Seems like a network problem.')
          engine.runAndWait()
          over()
          raise ValueError
        else:
          continue
    active()
    #scrolling down 
    driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
    time.sleep(2.5)
    
    timeout = time.time() + 120
    while True:
      try:  
        no_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="ISSUEREQ_OWNER"]')))
        ActionChains(driver).move_to_element(no_name_box).click(no_name_box).perform()
        time.sleep(0.5)
        no_name_box.send_keys(Keys.CONTROL,'a')
        time.sleep(0.1)
        ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.1)
        ActionChains(driver).send_keys(no_name).perform()
        #selecting NO name
        no_elem = '*//td[text()=' + '"' + no_name + '"' + ']'
        WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, no_elem))).click()
        break
      except (NoSuchElementException, TimeoutException):
        if  time.time() > timeout:
          engine.say('Please go back, Re fresh the page and try again. Seems like a network problem.')
          engine.runAndWait()
          over()
          raise ValueError
        else:
          continue
    #click on Save
    WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save"]'))).click()
    # time.sleep(0.4)
    # #click on Close to go back to case
    # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Close"]'))).click()
    # time.sleep(0.4)
    # #click on Edit to assign it to the DO
    # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//i[@class="icon icon-edit"]'))).click()
    # time.sleep(0.4)
    # #click on assessment
    # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
    # #click on sent back to DO or sent to DO
    # try:
    #   WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent back to DO"]'))).click()
    # except:
    #   WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent to DO"]'))).click()
      
    # timeout = time.time() + 120
    # while True:
    #   try:  
    #     #DO Name box
    #     do_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="cust_5788"]')))
    #     ActionChains(driver).move_to_element(do_name_box).click(do_name_box).perform()
    #     time.sleep(0.1)
    #     do_name_box.send_keys(Keys.CONTROL,'a')
    #     time.sleep(0.1)
    #     ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
    #     time.sleep(0.1)
    #     ActionChains(driver).send_keys(do_name).perform()
    #     #selecting DO name
    #     do_elem = '*//td[text()=' + '"' + do_name + '"' + ']'
    #     WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, do_elem))).click()
    #     break
    #   except (NoSuchElementException, TimeoutException):
    #     if  time.time() > timeout:
    #       engine.say('Re fresh the page and try again. Seems like a network problem.')
    #       engine.runAndWait()
    #       over()
    #       raise ValueError
    #     else:
    #       continue
    
    # timeout = time.time() + 120
    # while True:
    #   try:
    #     #delete previous value if any in complaint status
    #     prop_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="cust_6008"]')))
    #     prop_box.send_keys(Keys.CONTROL,'a')
    #     prop_box.send_keys(Keys.BACKSPACE)
    #     #send Complaint with bank value
    #     prop_box.send_keys(status)
    #     status_elem = '*//td[text()=' + '"' + status + '"' + ']'
    #     #wait for status to come and select
    #     WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, status_elem))).click()
    #     break
    #   except (NoSuchElementException, TimeoutException):
    #     if  time.time() > timeout:
    #       engine.say('Re fresh the page and try again. Seems like a network problem.')
    #       engine.runAndWait()
    #       over()
    #       raise ValueError
    #     else:
    #       continue
    #   # click on Save and Proceed
    # driver.find_element_by_xpath('*//span[text()="Save and Proceed"]').click()
    over()
  else:
    engine.say('For {}, please select Maintainable button.'.format(bank_in_cms))
    engine.runAndWait()
    over()
    raise ValueError

#mailer function
def mailer():
  active()
  cur = None
  cur1 = '9'
  office = data.loc['office'].to_dict()['value']
  try:
    c_92 = WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO - Closure Clause 9.2a"]')))
    cur = 'BO - Closure 9(2) (a)'
  except:
    try:
      c_93a = WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO -  Closure Clause 9.3a "]')))
      cur = 'BO - Closure 9(3) (a)'
    except:
      try:
        c_93b = WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO - Closure Clause 9.3b"]')))
        cur = 'BO - Closure 9(3) (b)'
      except:
        try:
          c_93c = WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO - Closure Clause 9.3c"]')))
          cur = 'BO 9(3)(C)'
        except:
          engine.say('Sorry, this clause is not yet supported')
          engine.runAndWait()
          over()
          raise ValueError

  win_all = len(driver.window_handles)
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Summary"]'))).click()
  email = WebDriverWait(driver, 0.15).until(EC.visibility_of_element_located((By.XPATH, '*//span[@data-autoid="cust_5349_ctrl"]'))).text
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Email Communications"]'))).click()
  WebDriverWait(driver, 6).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Compose Email"]'))).click()
     
  timeout = time.time() + 30
  
  while win_all == len(driver.window_handles):
    time.sleep(1)
    if  time.time() > timeout:
      engine.say('Re fresh the page and try again. Seems like a network problem.')
      engine.runAndWait()
      over()
      raise ValueError
      
  main_window_handle = None
  while not main_window_handle:
    main_window_handle = driver.current_window_handle
  
  email_handle = driver.window_handles[-1]
  driver.switch_to.window(email_handle)
  
  from_email = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//select[@name="EMAILFROMID"]')))
  ActionChains(driver).move_to_element(from_email).click().send_keys(office).send_keys(Keys.ENTER).send_keys(Keys.TAB).send_keys(email).perform()        
  WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="toggle__text"]'))).click()
  WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//a[@data-autoid="RelatedTemplateID_srch"]'))).click()
  srch = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="Grid_SearchTextBox"]')))
  ActionChains(driver).move_to_element(srch).click().send_keys(cur1).send_keys(Keys.ENTER).perform()
  time.sleep(1)
  clause_xpath = '*//div[@title=' + '"' + cur +  '"]'   
  WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, clause_xpath))).click()
  over()
  
  
def m113a():
  active()
  cur = None
  cur1 = '11'
  office = data.loc['office'].to_dict()['value']
  cur = 'BO - Closure 11(3) (a) Settlement'
  
  win_all = len(driver.window_handles)
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Summary"]'))).click()
  email = WebDriverWait(driver, 0.15).until(EC.visibility_of_element_located((By.XPATH, '*//span[@data-autoid="cust_5349_ctrl"]'))).text
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Email Communications"]'))).click()
  WebDriverWait(driver, 6).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Compose Email"]'))).click()
     
  timeout = time.time() + 30
  
  while win_all == len(driver.window_handles):
    time.sleep(1)
    if  time.time() > timeout:
      engine.say('Re fresh the page and try again. Seems like a network problem.')
      engine.runAndWait()
      over()
      raise ValueError
      
  main_window_handle = None
  while not main_window_handle:
    main_window_handle = driver.current_window_handle
  
  email_handle = driver.window_handles[-1]
  driver.switch_to.window(email_handle)
  
  from_email = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//select[@name="EMAILFROMID"]')))
  ActionChains(driver).move_to_element(from_email).click().send_keys(office).send_keys(Keys.ENTER).send_keys(Keys.TAB).send_keys(email).perform()        
  WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="toggle__text"]'))).click()
  WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//a[@data-autoid="RelatedTemplateID_srch"]'))).click()
  srch = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="Grid_SearchTextBox"]')))
  ActionChains(driver).move_to_element(srch).click().send_keys(cur1).send_keys(Keys.ENTER).perform()
  time.sleep(1)
  clause_xpath = '*//div[@title=' + '"' + cur +  '"]'   
  WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, clause_xpath))).click()
  over()
  
  
def m13a():
  active()
  cur = None
  cur1 = '13'
  office = data.loc['office'].to_dict()['value']
  cur = 'BO -13(1) (a) 13(1)(b) 13(1)(c) Non Appealable'
  
  win_all = len(driver.window_handles)
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Summary"]'))).click()
  email = WebDriverWait(driver, 0.15).until(EC.visibility_of_element_located((By.XPATH, '*//span[@data-autoid="cust_5349_ctrl"]'))).text
  WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Email Communications"]'))).click()
  WebDriverWait(driver, 6).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Compose Email"]'))).click()
     
  timeout = time.time() + 30
  
  while win_all == len(driver.window_handles):
    time.sleep(1)
    if  time.time() > timeout:
      engine.say('Re fresh the page and try again. Seems like a network problem.')
      engine.runAndWait()
      over()
      raise ValueError
      
  main_window_handle = None
  while not main_window_handle:
    main_window_handle = driver.current_window_handle
  
  email_handle = driver.window_handles[-1]
  driver.switch_to.window(email_handle)
  
  from_email = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//select[@name="EMAILFROMID"]')))
  ActionChains(driver).move_to_element(from_email).click().send_keys(office).send_keys(Keys.ENTER).send_keys(Keys.TAB).send_keys(email).perform()        
  WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="toggle__text"]'))).click()
  WebDriverWait(driver, 0.01).until(EC.visibility_of_element_located((By.XPATH, '*//a[@data-autoid="RelatedTemplateID_srch"]'))).click()
  srch = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="Grid_SearchTextBox"]')))
  ActionChains(driver).move_to_element(srch).click().send_keys(cur1).send_keys(Keys.ENTER).perform()
  time.sleep(1)
  clause_xpath = '*//div[@title=' + '"' + cur +  '"]'   
  WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, clause_xpath))).click()
  over()
  
def maintainableTWM_SBID(func):
  start()
  active()
  try:
    for i in range(len(driver.window_handles)):
      driver.switch_to.window(driver.window_handles[i+1])
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()
      get_list = func()
      bank_in_cms = driver.find_element_by_xpath('*//span[@data-autoid="cust_5958_ctrl"]').text
      no_name = get_list[0] 
      #checking status and storing the text in sts
      sts = WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//span[@data-autoid="CASE_STATUSCODE_ctrl"]'))).text
      if sts != 'New Complaint' or bank_in_cms != 'STATE BANK OF INDIA' :
        continue
      else:
        timeout = time.time() + 30
          #clicking on first edit button
        flg = True
        while flg == True:
          try:
            # finding edit menu hover icon and going to it
            edit_menu= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="action-hover"]')))
            ActionChains(driver).move_to_element(edit_menu).perform()
            #clicking on edit-icon
            edit_icon= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Edit"]')))
            driver.get(edit_icon.get_attribute('href'))
            flg = False
          except (NoSuchElementException, TimeoutException):
            if  time.time() > timeout:
              engine.say('Re fresh the page and try again. Seems like a network problem.')
              engine.runAndWait()
              over()
              raise ValueError
            else:
              continue
        driver.switch_to.window(driver.window_handles[i+1])
        #scrolling down 
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(2.5)
        
        timeout = time.time() + 120
        flg = True
        while flg == True:
          try:  
            no_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="ISSUEREQ_OWNER"]')))
            ActionChains(driver).move_to_element(no_name_box).click(no_name_box).perform()
            time.sleep(0.5)
            no_name_box.send_keys(Keys.CONTROL,'a')
            time.sleep(0.1)
            ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
            time.sleep(0.1)
            ActionChains(driver).send_keys(no_name).perform()
            #selecting NO name
            no_elem = '*//td[text()=' + '"' + no_name + '"' + ']'
            WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, no_elem))).click()
            flg = False
          except (NoSuchElementException, TimeoutException):
            if  time.time() > timeout:
              engine.say('Please go back, Re fresh the page and try again. Seems like a network problem.')
              engine.runAndWait()
              over()
              raise ValueError
            else:
              continue
        #click on Save
        WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save"]'))).click()
  except IndexError:
    over()  
    
def maintainableTWM_SBIK(func):
  start()
  active()
  try:
    for i in range(len(driver.window_handles)):
      driver.switch_to.window(driver.window_handles[i+1])
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()
      get_list = func()
      no_name = get_list[0]
      bank_in_cms = driver.find_element_by_xpath('*//span[@data-autoid="cust_5958_ctrl"]').text
      #checking status and storing the text in sts
      sts = WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//span[@data-autoid="CASE_STATUSCODE_ctrl"]'))).text
      if sts != 'New Complaint' or bank_in_cms != 'STATE BANK OF INDIA' :
        continue
      else:
        timeout = time.time() + 30
          #clicking on first edit button
        flg = True
        while flg == True:
          try:
            # finding edit menu hover icon and going to it
            edit_menu= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="action-hover"]')))
            ActionChains(driver).move_to_element(edit_menu).perform()
            #clicking on edit-icon
            edit_icon= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Edit"]')))
            driver.get(edit_icon.get_attribute('href'))
            flg = False
          except (NoSuchElementException, TimeoutException):
            if  time.time() > timeout:
              engine.say('Re fresh the page and try again. Seems like a network problem.')
              engine.runAndWait()
              over()
              raise ValueError
            else:
              continue
        driver.switch_to.window(driver.window_handles[i+1])
        #scrolling down 
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(2.5)
        
        timeout = time.time() + 120
        flg = True
        while flg == True:
          try:  
            no_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="ISSUEREQ_OWNER"]')))
            ActionChains(driver).move_to_element(no_name_box).click(no_name_box).perform()
            time.sleep(0.5)
            no_name_box.send_keys(Keys.CONTROL,'a')
            time.sleep(0.1)
            ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
            time.sleep(0.1)
            ActionChains(driver).send_keys(no_name).perform()
            #selecting NO name
            no_elem = '*//td[text()=' + '"' + no_name + '"' + ']'
            WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, no_elem))).click()
            flg = False
          except (NoSuchElementException, TimeoutException):
            if  time.time() > timeout:
              engine.say('Please go back, Re fresh the page and try again. Seems like a network problem.')
              engine.runAndWait()
              over()
              raise ValueError
            else:
              continue
        #click on Save
        WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save"]'))).click()

  except IndexError:
    over()  
  
#function for tab wise maintainable for BO Kanpur
def maintainableTWM():
  start()
  active()
  try:
    for i in range(len(driver.window_handles)):
      driver.switch_to.window(driver.window_handles[i+1])
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()
      
      bank_in_cms = driver.find_element_by_xpath('*//span[@data-autoid="cust_5958_ctrl"]').text
      no_name = data_banks.loc[bank_in_cms].to_dict()['value-NO']
      do_name = data_banks.loc[bank_in_cms].to_dict()['value-DO']
      status = 'Complaint with Bank'
      #checking status and storing the text in sts
      sts = WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//span[@data-autoid="CASE_STATUSCODE_ctrl"]'))).text
      if sts != 'New Complaint' or bank_in_cms == 'STATE BANK OF INDIA':
        continue
      else:
        timeout = time.time() + 30
          #clicking on first edit button
        flg = True
        while flg == True:
          try:
            # finding edit menu hover icon and going to it
            edit_menu= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//div[@class="action-hover"]')))
            ActionChains(driver).move_to_element(edit_menu).perform()
            #clicking on edit-icon
            edit_icon= WebDriverWait(driver, 12).until(EC.visibility_of_element_located((By.XPATH, '*//a[@title="Edit"]')))
            driver.get(edit_icon.get_attribute('href'))
            flg = False
          except (NoSuchElementException, TimeoutException):
            if  time.time() > timeout:
              engine.say('Re fresh the page and try again. Seems like a network problem.')
              engine.runAndWait()
              over()
              raise ValueError
            else:
              continue
        driver.switch_to.window(driver.window_handles[i+1])
        #scrolling down 
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(2.5)
        
        timeout = time.time() + 120
        flg = True
        while flg == True:
          try:  
            no_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="ISSUEREQ_OWNER"]')))
            ActionChains(driver).move_to_element(no_name_box).click(no_name_box).perform()
            time.sleep(0.5)
            no_name_box.send_keys(Keys.CONTROL,'a')
            time.sleep(0.1)
            ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
            time.sleep(0.1)
            ActionChains(driver).send_keys(no_name).perform()
            #selecting NO name
            no_elem = '*//td[text()=' + '"' + no_name + '"' + ']'
            WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, no_elem))).click()
            flg = False
          except (NoSuchElementException, TimeoutException):
            if  time.time() > timeout:
              engine.say('Please go back, Re fresh the page and try again. Seems like a network problem.')
              engine.runAndWait()
              over()
              raise ValueError
            else:
              continue
        #click on Save
        WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save"]'))).click()
        # time.sleep(0.4)
        # #click on Close to go back to case
        # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Close"]'))).click()
        # time.sleep(0.4)
        # #click on Edit to assign it to the DO
        # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//i[@class="icon icon-edit"]'))).click()
        # time.sleep(0.4)
        # #click on assessment
        # WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
        # #click on sent back to DO or sent to DO
        # try:
        #   WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent back to DO"]'))).click()
        # except:
        #   WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Sent to DO"]'))).click()
          
        # timeout = time.time() + 120
        # flg = True
        # while flg == True:
        #   try:  
        #     #DO Name box
        #     do_name_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="cust_5788"]')))
        #     ActionChains(driver).move_to_element(do_name_box).click(do_name_box).perform()
        #     time.sleep(0.1)
        #     do_name_box.send_keys(Keys.CONTROL,'a')
        #     time.sleep(0.1)
        #     ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
        #     time.sleep(0.1)
        #     ActionChains(driver).send_keys(do_name).perform()
        #     #selecting DO name
        #     do_elem = '*//td[text()=' + '"' + do_name + '"' + ']'
        #     WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, do_elem))).click()
        #     flg = False
          # except (NoSuchElementException, TimeoutException):
          #   if  time.time() > timeout:
          #     engine.say('Re fresh the page and try again. Seems like a network problem.')
          #     engine.runAndWait()
          #     over()
          #     raise ValueError
          #   else:
          #     continue
        
        # timeout = time.time() + 120
        # flg = True
        # while flg == True:
        #   try:
        #     #delete previous value if any in complaint status
        #     prop_box = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//input[@name="cust_6008"]')))
        #     prop_box.send_keys(Keys.CONTROL,'a')
        #     prop_box.send_keys(Keys.BACKSPACE)
        #     #send Complaint with bank value
        #     prop_box.send_keys(status)
        #     status_elem = '*//td[text()=' + '"' + status + '"' + ']'
        #     #wait for status to come and select
        #     WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, status_elem))).click()
        #     flg = False
        #   except (NoSuchElementException, TimeoutException):
        #     if  time.time() > timeout:
        #       engine.say('Re fresh the page and try again. Seems like a network problem.')
        #       engine.runAndWait()
        #       over()
        #       raise ValueError
        #     else:
        #       continue
        #   # click on Save and Proceed
        # driver.find_element_by_xpath('*//span[text()="Save and Proceed"]').click()
  except IndexError:
    over()
    
# Function to get all contents on the dashboard to an excel

def to_xl():
  start()
  active()
  #finding the len of all cases on page
  records = driver.find_elements_by_xpath('//div[@class="showrecords"]')[0].text
  records = records.split()
  hyphen_index = records[1].index('-')
  records = records[1]
  first = int(records[:hyphen_index])
  second = int(records[hyphen_index+1:])
  cases_count = second-first+1
  #getting xpath of all [0][0] . [0][1] etc elements
  all_rows = []
  
  try:
    for i in range(cases_count):
      all_divs = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' + str(i+1) + ']/div[1]/div[' + str(j) + ']' for j in range(1,11)]
      all_rows.append(all_divs)
      
    final_div_data = []
    
    #converting all these elements to text
    for i in range(cases_count):
      temp = []
      temp = [driver.find_element_by_xpath(all_rows[i][j]).text for j in range(10)]
      final_div_data.append(temp)
    
    df = pd.DataFrame(final_div_data, columns = ['Complaint No', 'Name', 'Entity', 'Prop Clause', 'Comp Cat', 'Comp Status', 'Reviewer', 'DO', 'Created On', ' '])
    
    local_time = time.ctime(time.time())
    local_time = local_time.replace(':','_')
    df = df.set_index('Complaint No')
    
    writer = pd.ExcelWriter("{}.xlsx".format(local_time), engine = 'xlsxwriter')
    df.to_excel(writer, sheet_name="TOTAL")
  
    workbook = writer.book
    worksheet = writer.sheets["TOTAL"]
    
    format1 = workbook.add_format({'num_format':'0'})
    format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
    format4 = workbook.add_format({'bold': True})
    
    worksheet.set_column('A:A', 24, format1)
    worksheet.set_column('B:B', 17, None)
    worksheet.set_column('C:C', 30, None)
    worksheet.set_column('D:D', 20, None)
    worksheet.set_column('E:E', 20, None)
    worksheet.set_column('F:F', 18, format2)
    worksheet.set_column('G:G', 10, None)
    worksheet.set_column('H:H', 10, None)
    worksheet.set_column('I:I', 14, format4)
    worksheet.set_column('J:J', 14, None)
    
    writer.save()
    writer.close()
    over()      
  except:
    for i in range(cases_count):
      all_divs = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' + str(i+1) + ']/div[1]/div[' + str(j) + ']' for j in range(1,7)]
      all_rows.append(all_divs)
      
    final_div_data = []
    
    #converting all these elements to text
    for i in range(cases_count):
      temp = []
      temp = [driver.find_element_by_xpath(all_rows[i][j]).text for j in range(6)]
      final_div_data.append(temp)
    
    df = pd.DataFrame(final_div_data, columns = ['Complaint No', 'Process', 'Comp Name', 'Bank Name', 'Fwd By', 'Created On'])
    
    local_time = time.ctime(time.time())
    local_time = local_time.replace(':','_')
    df = df.set_index('Complaint No')
    
    writer = pd.ExcelWriter("{}.xlsx".format(local_time), engine = 'xlsxwriter')
    df.to_excel(writer, sheet_name="TOTAL")
  
    workbook = writer.book
    worksheet = writer.sheets["TOTAL"]
    
    format1 = workbook.add_format({'num_format':'0'})
    format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
    format4 = workbook.add_format({'bold': True})
    
    worksheet.set_column('A:A', 24, format1)
    worksheet.set_column('B:B', 17, None)
    worksheet.set_column('C:C', 30, None)
    worksheet.set_column('D:D', 20, None)
    worksheet.set_column('E:E', 20, None)
    worksheet.set_column('F:F', 18, format2)
    
    writer.save()
    writer.close()
    over()
  
# Function to approve 11(3)(a) Tab Wise

def BO_approve_set():
  start()
  active()
  clause_text = data.loc['B113a'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #Selecting BO Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO Decision"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #select proposal box - M/NM and send 11(3)(a) value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('m')
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            time.sleep(0.1)
            # status_box.send_keys(Keys.BACKSPACE)
            # time.sleep(0.2)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()             
            # clause_box.send_keys(Keys.CONTROL + 'a')
            time.sleep(0.1)
            # clause_box.send_keys(Keys.BACKSPACE)
            # time.sleep(0.2)
            clause_box.clear()
            clause_box.send_keys('11 (3) (a)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break
          else:
            #select proposal box - M/NM and send 11(3)(a) value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('m')
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            time.sleep(0.1)
            do_box.clear()
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.3)
            # do_box.send_keys(Keys.BACKSPACE)
            # time.sleep(0.3)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('11 (3) (a)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

# Function to approve 13(a) Tab Wise

def BO_approve_rej():
  start()
  active()
  clause_text = data.loc['B13a'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #Selecting BO Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO Decision"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #select proposal box - M/NM and send 13-a value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('m')
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('13 (1) (a)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break
          else:
            #select proposal box - M/NM and send 13-a value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('m')
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # time.sleep(0.3)
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.3)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('13 (1) (a)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()
  
def BO_approve_92():  
  start()
  active()
  clause_text = data.loc['B92'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #Selecting BO Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO Decision"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #select proposal box - M/NM and send 9-2 value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('9 (2)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break
          else:
            #select proposal box - M/NM and send 9-2 value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # time.sleep(0.3)
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.3)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('9 (2)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def BO_approve_93c():  
  start()
  active()
  clause_text = data.loc['B93c'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #Selecting BO Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO Decision"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #select proposal box - M/NM and send 9-2 value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('9 (3) (c)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break
          else:
            #select proposal box - M/NM and send 9-2 value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # time.sleep(0.3)
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.3)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('9 (3) (c)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def BO_approve_35():  
  start()
  active()
  clause_text = data.loc['B35'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #Selecting BO Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO Decision"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #select proposal box - M/NM and send 9-2 value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('3 (5)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break
          else:
            #select proposal box - M/NM and send 9-2 value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # time.sleep(0.3)
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.3)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('3 (5)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def BO_approve_93a():  
  start()
  active()
  clause_text = data.loc['B93a'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      #Selecting BO Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO Decision"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #select proposal box - M/NM and send 9-3-a value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #Send Yes in FRC
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[2]/div/div/div/div/div/select'))).send_keys('y')
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('9 (3) (a)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break
          else:
            #select proposal box - M/NM and send 9-3-a value
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/select'))).send_keys('n')
            #Send Yes in FRC
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[2]/div/div/div/div/div/select'))).send_keys('y')
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # time.sleep(0.3)
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.3)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in complaint status and sending status
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')            
            #wait for status to come and select 
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('9 (3) (a)')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()
  
#Function to issue advisory and load text from excel file BIssueAdvisory
def IssueAdvisory():
  start()
  active()
  clause_text = data.loc['BIssueAdvisory'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then BO Decision make it maintainable
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="BO Decision"]'))).click()
      #now send m value to make it maintainable
      clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", clause_box)
      ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()     
      clause_box.send_keys('m')      
      #click on Final Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Final Decision"]'))).click()
      #Selecting Issue Advisory
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Issue Advisory"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            #send monetary to Advisory Type box
            ad_type_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select')))
            ad_type_box.send_keys('m')
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('Issue Advisory')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break
          else:
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()            
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            #send monetary to Advisory Type box
            ad_type_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/select')))
            ad_type_box.send_keys('m')
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('Issue Advisory')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[4]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).key_down(Keys.TAB).key_up(Keys.TAB).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()  

def NBFCO_IssueAdvisory():
  start()
  active()
  clause_text = data.loc['NBFCIssueAdvisory'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then BO Decision make it maintainable
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      #now send m value to make it maintainable
      clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", clause_box)
      ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()     
      clause_box.send_keys('m')      
      #click on Final Decision
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Final Decision"]'))).click()
      #Selecting Issue Advisory
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Issue Advisory"]'))).click()
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('Issue Advisory')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            ActionChains(driver).send_keys(Keys.TAB).perform()
            break
          else:
            #delete in DO field to enter DO name
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", do_box)
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            #delete previous value if any in proposed clause and sending clause
            clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", clause_box)
            ActionChains(driver).move_to_element(clause_box).click(clause_box).perform()            
            # clause_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # clause_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            clause_box.clear()
            clause_box.send_keys('Issue Advisory')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/input')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()            
            # status_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # status_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()    

            ActionChains(driver).send_keys(Keys.TAB).perform()
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      #sending TAB and values in comments field
      ActionChains(driver).send_keys(clause_text).perform()  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()  

def NBFCO_114a():
  start()
  active()
  clause_text = data.loc['NBFC11.4.a'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')      
      # now send m value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('m')
      # send m key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('m')
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('11.4.a')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('11.4.a')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('11.4.a')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('11.4.a')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()  

def NBFCO_131a():
  start()
  active()
  clause_text = data.loc['NBFC13.1a'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')      
      # now send m value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('m')
      # send m key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('m')
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('13.1a')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('13.1.a')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('13.1a')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('13.1.a')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()   

def NBFCO_132():
  start()
  active()
  clause_text = data.loc['NBFC13.2'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')      
      # now send m value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('m')
      # send m key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('m')
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('13.2')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('13.2')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('13.2')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('13.2')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def NBFCO_91a():
  start()
  active()
  clause_text = data.loc['NBFC9.1.A'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')
      # now send n value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('n')
      # send n key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('n')      
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.1.A.')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9.1.A')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.1.A.')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9.1.A')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def NBFCO_9Aa():
  start()
  active()
  clause_text = data.loc['NBFC9Aa'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send Yes in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('y')
      # now send n value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('n')
      # send n key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('n')      
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.A.a')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9-A.a')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.A.a')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9-A.a')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def NBFCO_9ac():
  start()
  active()
  clause_text = data.loc['NBFC9.A.c'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')      
      # now send n value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('n')
      # send n key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('n')      
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.A.c')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9-A.c')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.A.c')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9-A.c')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def NBFCO_91():
  start()
  active()
  clause_text = data.loc['NBFC91'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')      
      # now send n value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('n')
      # send n key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('n')      
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.1 Not on')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9.1')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr[1]/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9.1 Not on')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9.1')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr[1]/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def NBFCO_9Ag():
  start()
  active()
  clause_text = data.loc['NBFC9.A.g'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')      
      # now send n value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('n')
      # send n key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('n')      
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9-A.g')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9-A.g')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr[1]/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 9-A.g')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('9-A.g')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr[1]/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

def tick_all():
  start()
  case_nos = entryS.get().split()
  active()
  #finding the len of all cases on page
  records = driver.find_elements_by_xpath('//div[@class="showrecords"]')[0].text
  records = records.split()
  hyphen_index = records[1].index('-')
  records = records[1]
  first = int(records[:hyphen_index])
  second = int(records[hyphen_index+1:])
  cases_count = second-first+1
  #checking for type of check box and selecting that value
  for i in range(50):
    try:
      driver.find_element_by_xpath('//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[1]/div[1]/div[' + str(i+1) +']/div/div/span/div/label')
      divnumber = str(i+1)
      break
    except:
      None
  #getting xpath & elements for all check boxes
  all_checks_xpath = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' + str(i+1) + ']/div[1]/div[' + divnumber +']/div/div/span/div/label' for i in range(cases_count)]
  all_checks_elems = [driver.find_element_by_xpath(all_checks_xpath[i]) for i in range(cases_count)] 
  #getting case nos for all cases on page
  all_cases_xpath = ['*//a[@data-autoid="CAS_EX1_114_' + str(i) + '"]' for i in range(cases_count)]
  case_nos_elems = [driver.find_element_by_xpath(all_cases_xpath[i]) for i in range(cases_count)]
  case_nos_text = [k.get_attribute('title') for k in case_nos_elems]
  #match cases in entry to ticks and ticking them if found
  matching_indices = []
  for i in case_nos:
    if i in case_nos_text:
      matching_indices.append(case_nos_text.index(i))
  #ticking matched idices corresponding tick boxes
  for index in matching_indices:
    all_checks_xpath = ['//*[@id="objectWrapper"]/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' + str(i+1) + ']/div[1]/div[' + divnumber +']/div/div/span/div/label' for i in range(cases_count)]
    all_checks_elems = [driver.find_element_by_xpath(all_checks_xpath[i]) for i in range(cases_count)]    
    all_checks_elems[index].click()
  over()

def NBFCO_35():
  start()
  active()
  clause_text = data.loc['NBFC35'].to_dict()['value']
  bo_sec = data.loc['BO Secretary'].to_dict()['value']
  mult_bosec = False
  if ',' in bo_sec:
    bo_sec1,bo_sec2 = bo_sec.split(',')
    mult_bosec = True
  j=0
  for i in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[i+1])
      #click on Assessment then NBFCO Decision 
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Assessment"]'))).click()
      WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="NBFCO Decision"]'))).click()
      # Send No in FRC Box
      frc = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5389_ctrl"]')))
      driver.execute_script("arguments[0].click();", frc)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      frc.send_keys('n')
      # now send n value to NBFCO Decision field make it maintainable
      nbfco_d = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5388_ctrl"]')))
      driver.execute_script("arguments[0].click();", nbfco_d)
      # ActionChains(driver).move_to_element(nbfco_d).click(nbfco_d).perform()     
      nbfco_d.send_keys('n')
      # send n key to Proposed action box    
      prop_action_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//select[@data-autoid="cust_5939_ctrl"]')))
      driver.execute_script("arguments[0].click();", prop_action_box)
      # ActionChains(driver).move_to_element(prop_action_box).click(prop_action_box).perform()     
      prop_action_box.send_keys('n')      
      # send clause to clause
      timeout = time.time() + 200
      #get DO's name in CMS
      dofound = False
      try:
        putup_by_elems = driver.find_elements_by_xpath('*//div[@class="discussionThread__title"]')
        putup_by_list = [k.text for k in putup_by_elems]
        for i in range(len(putup_by_list)):
          if mult_bosec == True:
            if putup_by_list[i] != bo_sec1 and putup_by_list[i] != bo_sec2:
              do = putup_by_list[i]
              dofound = True
              break            
          else:
            if putup_by_list[i] != bo_sec:
              do = putup_by_list[i]
              dofound = True
              break
      except:
        dofound = False
      while True:
        try:
          if dofound == False:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 3.5')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('3.5')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break
          else:
            #delete previous value if any in proposed clause and sending clause
            prop_clause_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5946_ctrl"]')))
            driver.execute_script("arguments[0].click();", prop_clause_box)
            ActionChains(driver).move_to_element(prop_clause_box).click(prop_clause_box).perform()            
            time.sleep(0.1)
            prop_clause_box.clear()
            prop_clause_box.send_keys('Clause 3.5')
            #wait for clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div/table/tbody/tr/td[1]'))).click()            
            #select status box - send status value
            status_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_6008_ctrl"]')))
            driver.execute_script("arguments[0].click();", status_box)
            ActionChains(driver).move_to_element(status_box).click(status_box).perform()             
            time.sleep(0.1)
            status_box.clear()
            status_box.send_keys('Complaint in Process')
            #wait for status to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            # Send closure clause
            closure_clause = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5673_ctrl"]')))
            driver.execute_script("arguments[0].click();", closure_clause)
            ActionChains(driver).move_to_element(closure_clause).click(closure_clause).perform()             
            time.sleep(0.1)
            closure_clause.clear()
            closure_clause.send_keys('3.5')
            #wait for closure clause to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[2]/div/div/div/div/div/div/table/tbody/tr/td'))).click()
            # Send DO Name and select
            do_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5788_ctrl"]')))
            driver.execute_script("arguments[0].click();", do_box)
            # ActionChains(driver).move_to_element(do_box).click(do_box).perform()            
            # do_box.send_keys(Keys.CONTROL + 'a')
            # time.sleep(0.2)
            # do_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            do_box.clear()
            ActionChains(driver).move_to_element(do_box).click(do_box).perform()
            do_box.send_keys(do)  
            #wait for do name to come and select
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[1]/div/div/div[2]/div[1]/div/div/div[5]/div[1]/div/div/div/div/div/div/table/tbody/tr/td'))).click()            
            break            
        except (NoSuchElementException, TimeoutException):
          if  time.time() > timeout:
            engine.say('Re fresh the page and try again. Seems like a network problem.')
            engine.runAndWait()
            over()
            raise ValueError
          else:
            continue
      
      comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5339_ctrl"]')))
      driver.execute_script("arguments[0].click();", comments_box)
      # ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()        
      time.sleep(0.1)
      comments_box.clear()
      ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()
      comments_box.send_keys(clause_text)  
      #click on Save and Proceed
      # WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save and Proceed"]'))).click()
    except (NoSuchElementException, TimeoutException, IndexError):
      j += 1
      if j >=2:
        break  
  over()

#Function to search cases then go to NO Record then change from Info Required to 10-1 Notice Issued

def to101_load():
  start()
  case_nos = entryS.get().split()
  comments_text = data.loc['Comments10_1'].to_dict()['value']
  
  active()
  for i in range(len(case_nos)):
    driver.execute_script('''window.open("about:blank",'_newtab_home');''')
    driver.switch_to.window(driver.window_handles[-1])
    driver.get('https://cms.rbi.org.in/ro/app/CRMNextObject/SearchAction/Case')
    in_field = driver.find_element_by_xpath('//*[@id="objectWrapper"]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div[1]/div[1]/div/div/div/input')
    in_field.send_keys(case_nos[i] + Keys.ENTER)
    try:
      WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()
    except:
      engine.say('I have waited for more than 40 seconds, please check your internet connection !')
      engine.runAndWait()
      break
    #finding the len of all cases on page
    records = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="showrecords"]'))).text
    records = records.split()
    hyphen_index = records[1].index('-')
    records = records[1]
    first = int(records[:hyphen_index])
    second = int(records[hyphen_index+1:])
    cases_count = second-first+1
    #getting xpath & elements for all NO Record Statuses
    all_NOR_xpath = ['*//div[@data-autoid="StatusCode_' + str(i) + '"]' for i in range(cases_count)]
    all_NOR_text = [driver.find_element_by_xpath(all_NOR_xpath[i]).text for i in range(cases_count)]
    
    #getting edit button elems for all NO Records on page 
    xpath_links_all = ['//*[@id="newobject"]/div[1]/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div/div/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[2]' for x in range(cases_count)]
    xpath_edit_menu = ['//*[@id="newobject"]/div[1]/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div/div/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[2]/div/div[2]/div/a' for x in range(cases_count)] 
    
    #match and get indices with those having Information Required
    matching_indices = []
    for i in all_NOR_text:
      if i == 'Information Required':
        matching_indices.append(all_NOR_text.index(i))  
    
    #counter to have matching cases length
    counter = len(matching_indices)
    #only do when found matching records
    if counter != 0:
      #to loop through each cycle of Info Reqd complaint
      all_done = False
      while all_done == False:        
        for n in matching_indices:
          WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()          
          x = True
          while x:
            edit_menu= WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, xpath_links_all[n])))
            ActionChains(driver).move_to_element(edit_menu).perform()
            try:
              temp = driver.find_element_by_xpath(xpath_edit_menu[n])
              ActionChains(driver).move_to_element(temp).click().perform()
              x = False
            except NoSuchElementException:
              continue
          status_code_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div[1]/div[2]/div/div/div/div/div/select')))
          status_code_box.send_keys('10')
          
          #update date to 5 days from now
          t = datetime.date.today()
          t += timedelta(5)
          t = t.strftime("%d/%m/%Y")
          # update this calculated date
          date_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5919_ctrl"]')))
          driver.execute_script("arguments[0].click();", date_box)
          ActionChains(driver).move_to_element(date_box).click(date_box).perform()     
          date_box.send_keys(t)          
          
          
          #enter comments to NO
          comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5916_ctrl"]')))
          driver.execute_script("arguments[0].click();", comments_box)
          ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()     
          comments_box.send_keys(comments_text)
          
          #click on Save
          WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save"]'))).click()
          
          #decrease counter by 1 and exit if it reaches 0 or else click on close button for others to proceed
          counter -= 1
          if counter == 0:
            all_done = True
          else:
            time.sleep(0.3)
            # #click on Close to go back to case
            WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Close"]'))).click()
            time.sleep(0.3)          
  over() 
  
#Function to go to NO Record then change from Info Required to 10-1 Notice Issued(on already loaded tabs)
  
def to101_loaded():
  start()
  active()
  comments_text = data.loc['Comments10_1'].to_dict()['value']  
  for t in range(len(driver.window_handles)):
    try:
      driver.switch_to.window(driver.window_handles[t+1])
      try:
        WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()
      except:
        engine.say('I have waited for more than 40 seconds, please check your internet connection !')
        engine.runAndWait()
        break
      #finding the len of all cases on page
      records = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="showrecords"]'))).text
      records = records.split()
      hyphen_index = records[1].index('-')
      records = records[1]
      first = int(records[:hyphen_index])
      second = int(records[hyphen_index+1:])
      cases_count = second-first+1
      #getting xpath & elements for all NO Record Statuses
      all_NOR_xpath = ['*//div[@data-autoid="StatusCode_' + str(i) + '"]' for i in range(cases_count)]
      all_NOR_text = [driver.find_element_by_xpath(all_NOR_xpath[i]).text for i in range(cases_count)]
      
      #getting edit button elems for all NO Records on page 
      xpath_links_all = ['//*[@id="newobject"]/div[1]/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div/div/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[2]' for x in range(cases_count)]
      xpath_edit_menu = ['//*[@id="newobject"]/div[1]/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div/div/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[' +str(x+1)+ ']/div[2]/div[2]/div/div[2]/div/a' for x in range(cases_count)] 
      
      #match and get indices with those having Information Required
      matching_indices = []
      for i in all_NOR_text:
        if i == 'Information Required':
          matching_indices.append(all_NOR_text.index(i))  
      
      #counter to have matching cases length
      counter = len(matching_indices)
      #only do when found matching records
      if counter != 0:
        #to loop through each cycle of Info Reqd complaint
        all_done = False
        while all_done == False:        
          for n in matching_indices:
            WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Nodal Officer Record"]'))).click()          
            x = True
            while x:
              edit_menu= WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, xpath_links_all[n])))
              ActionChains(driver).move_to_element(edit_menu).perform()
              try:
                temp = driver.find_element_by_xpath(xpath_edit_menu[n])
                ActionChains(driver).move_to_element(temp).click().perform()
                x = False
              except NoSuchElementException:
                continue
            status_code_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="newobject"]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div[1]/div[2]/div/div/div/div/div/select')))
            status_code_box.send_keys('10')
            
            #update date to 5 days from now
            t = datetime.date.today()
            t += timedelta(5)
            t = t.strftime("%d/%m/%Y")
            # update this calculated date
            date_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//input[@data-autoid="cust_5919_ctrl"]')))
            driver.execute_script("arguments[0].click();", date_box)
            ActionChains(driver).move_to_element(date_box).click(date_box).perform()     
            date_box.send_keys(t)              
            
            #enter comments to NO
            comments_box = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '*//textarea[@data-autoid="cust_5916_ctrl"]')))
            driver.execute_script("arguments[0].click();", comments_box)
            ActionChains(driver).move_to_element(comments_box).click(comments_box).perform()     
            comments_box.send_keys(comments_text)
            
            #click on Save
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Save"]'))).click()
            
            #decrease counter by 1 and exit if it reaches 0 or else click on close button for others to proceed
            counter -= 1
            if counter == 0:
              all_done = True
            else:
              time.sleep(0.3)
              # #click on Close to go back to case
              WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, '*//span[text()="Close"]'))).click()
              time.sleep(0.3)  
    except Exception as e: 
      print(e)
      break
  over() 
  
  
#generator to yield x,y values for button positioning
def x_y(x,y):
  while x <= 965:
    x += 305
    yield x
    yield y
    if x == 965:
      x = 50
      y += 50
      continue


#initializing generator
a = x_y(x=50,y=120)


window = ThemedTk(theme = 'radiance')
window.title('The Real CMS')
window.geometry("1200x650")
window.maxsize(width=1300, height=700)
canvas = Canvas(window, width = 1200, height= 650)
canvas.create_rectangle(330, 115, 1190, 265, fill = 'dark grey')
canvas.create_oval(30, 185, 70, 215, fill = 'dark grey')
canvas.create_text(90,209,text='Common Actions', anchor='sw', font= ('Georgia',10,'bold'))
canvas.create_rectangle(330, 265, 1190, 415, fill = 'light pink')
canvas.create_rectangle(867, 415, 1190, 465, fill = 'light pink')
canvas.create_oval(30, 315, 70, 345, fill = 'light pink')
canvas.create_text(90,340,text='Processing Team Actions', anchor='sw', font= ('Georgia',10,'bold'))
canvas.create_rectangle(330, 415, 867, 465, fill = 'sky blue')
canvas.create_oval(30, 415, 70, 455, fill = 'sky blue')
canvas.create_text(90,440,text='Secretary Actions', anchor='sw', font= ('Georgia',10,'bold'))
canvas.create_text(830,485,text='(Maintainable buttons for SBI having two NOs (TWM & Single Run))', anchor='sw', font= ('Arial Narrow',8,'bold'))
canvas.grid(row=0,column=0)

# window.configure(bg='light pink')
style = ttk.Style()
style.map("C.TButton",
    foreground=[('pressed', 'red'), ('active', 'blue')],
    background=[('pressed', '!disabled', 'black'), ('active', 'white')]
    )

ttk.Style().configure("C.TButton", padding=0, relief="flat", width = 18)
buttonQ = tk.Button(canvas, text='Update Values from Excel', command= excel_updater, bg='lightpink',fg='blue')
buttonQ.grid(row=5,column=5)
buttonQ.place(x=975,y=570)
buttonQ = tk.Button(canvas, text='EXIT CMS Manager', command= lambda:[driver.quit(), window.destroy()], bg='red',fg='white')
buttonQ.grid(row=5,column=5)
buttonQ.place(x=1000,y=611)
buttonLI = tk.Button(canvas, text = 'Login to CMS', command = lambda:[start(),CMS_login()], bg='green', fg='white', font = "Helvetica 10 bold")
buttonLI.grid(row=5,column=5)
buttonLI.place(x=50, y=40)
buttonLI = tk.Button(canvas, text = 'Export Dash to Excel', command = to_xl, bg='lightpink', fg='black', font = "Helvetica 10 bold")
buttonLI.grid(row=5,column=5)
buttonLI.place(x=50, y=100)
buttonLO = tk.Button(canvas, text = 'Logout of CMS', bg='red', fg='white', font = "Helvetica 10 bold", command=logout)
buttonLO.grid(row=5,column=5)
buttonLO.place(x=1050, y=40)
buttonSD = ttk.Button(canvas, text = 'Edit SFA Drafts', style="C.TButton", command = open_drafts_SFA)
buttonSD.grid(row=5,column=5)
buttonSD.place(x=next(a), y=next(a))
buttonNC = ttk.Button(canvas, text = 'Open New Complaints', style="C.TButton", command = open_New_C)
buttonNC.grid(row=5,column=5)
buttonNC.place(x=next(a), y=next(a))
buttonNC = ttk.Button(canvas, text = 'Edit BO Decisions', style="C.TButton", command = open_BO_D)
buttonNC.grid(row=5,column=5)
buttonNC.place(x=next(a), y=next(a))
buttonOA = ttk.Button(canvas, text = 'Open All On Page', style="C.TButton", command = open_all)
buttonOA.grid(row=5,column=5)
buttonOA.place(x=next(a), y=next(a))
buttonRA = ttk.Button(canvas, text = 'Refresh All Tabs', style="C.TButton", command = refresh_all)
buttonRA.grid(row=5,column=5)
buttonRA.place(x=next(a), y=next(a))
buttonEA = ttk.Button(canvas, text = 'Edit All On Page', style="C.TButton", command = edit_all)
buttonEA.grid(row=5,column=5)
buttonEA.place(x=next(a), y=next(a))
buttonMl = ttk.Button(canvas, text = 'Mailer-9(2,3.a,3.b,3.c)', style="C.TButton",command = mailer)
buttonMl.grid(row=5,column=5)
buttonMl.place(x=next(a), y=next(a))
buttonCASFA = ttk.Button(canvas, text = 'Close SFA as NAC', style="C.TButton",command = CASFA)
buttonCASFA.grid(row=5,column=5)
buttonCASFA.place(x=next(a), y=next(a))
buttonactive = ttk.Button(canvas, text = 'Switch to Active Tab', style="C.TButton",command = active)
buttonactive.grid(row=5,column=5)
buttonactive.place(x=next(a), y=next(a))
a = x_y(x=50,y=270)
button8182 = ttk.Button(canvas, text = '8(1)8(2) to BO', style="C.TButton", command = lambda:[NMC(_8182)])
button8182.grid(row=5,column=5)
button8182.place(x=next(a), y=next(a))
button91 = ttk.Button(canvas, text = '9(1) to BO', style="C.TButton", command = lambda:[NMC(_91)])
button91.grid(row=5,column=5)
button91.place(x=next(a), y=next(a))
button92 = ttk.Button(canvas, text = '9(2) to BO', style="C.TButton", command = lambda:[NMC(_92)])
button92.grid(row=5,column=5)
button92.place(x=next(a), y=next(a))
button93a = ttk.Button(canvas, text = '9(3)(a) to BO', style="C.TButton", command = lambda:[NMC(_93a)])
button93a.grid(row=5,column=5)
button93a.place(x=next(a), y=next(a))
button93b = ttk.Button(canvas, text = '9(3)(b) to BO', style="C.TButton", command = lambda:[NMC(_93b)])
button93b.grid(row=5,column=5)
button93b.place(x=next(a), y=next(a))
button7191b = ttk.Button(canvas, text = '7(1)9(1)b to BO', style="C.TButton", command = lambda:[NMC(_7191b)])
button7191b.grid(row=5,column=5)
button7191b.place(x=next(a), y=next(a))
button7191c = ttk.Button(canvas, text = '7(1)9(1)cc to BO', style="C.TButton", command = lambda:[NMC(_7191c)])
button7191c.grid(row=5,column=5)
button7191c.place(x=next(a), y=next(a))
button93c = ttk.Button(canvas, text = '9(3)(c) to BO', style="C.TButton", command = lambda:[NMC(_93c)])
button93c.grid(row=5,column=5)
button93c.place(x=next(a), y=next(a))
buttonM = ttk.Button(canvas, text = 'MAINTAINABLE', style="C.TButton",command = maintainable)
buttonM.grid(row=5,column=5)
buttonM.place(x=next(a), y=next(a))
buttonM113a = tk.Button(canvas, text = '11.3.a',fg='purple',bg='silver', command= m113a)
buttonM113a.grid(row=5,column=5)
buttonM113a.place(x=555, y=202)
buttonM13a = tk.Button(canvas, text = '13.a',fg='purple',bg='silver', command = m13a)
buttonM13a.grid(row=5,column=5)
buttonM13a.place(x=562, y=238)
a = x_y(x=50,y=420)
button113A = ttk.Button(canvas, text = '11(3)(a) to BO', style="C.TButton", command = resolved)
button113A.grid(row=5,column=5)
button113A.place(x=next(a), y=next(a))
button13A = ttk.Button(canvas, text = '11(3)(a) to BO - SP', style="C.TButton", command = resolved_SP)
button13A.grid(row=5,column=5)
button13A.place(x=next(a), y=next(a))
#tab wise maintinable button
buttonTWM = tk.Button(canvas, text = 'TWM',fg='black',bg='cyan', command = maintainableTWM)
buttonTWM.grid(row=5,column=5)
buttonTWM.place(x=900, y=370)
#placing SBI buttons TWM and Single Run
buttonIBK = tk.Button(canvas, text = 'SBI-D',fg='black',bg='cyan', command = lambda:[maintainableTWM_SBID(SBIDEL)])
buttonIBK.grid(row=5,column=5)
buttonIBK.place(x=880, y=425)
buttonIBN = tk.Button(canvas, text = 'SBI-K',fg='black',bg='cyan', command = lambda:[maintainableTWM_SBID(SBICHD)])
buttonIBN.grid(row=5,column=5)
buttonIBN.place(x=955, y=425)
buttonSBID = tk.Button(canvas, text = 'SBI-D',fg='purple',bg='silver', command =lambda:[maintainable_bank(SBIDEL)])
buttonSBID.grid(row=5,column=5)
buttonSBID.place(x=1030, y=425)
buttonSBIC = tk.Button(canvas, text = 'SBI-K',fg='purple',bg='silver', command =lambda:[maintainable_bank(SBICHD)])
buttonSBIC.grid(row=5,column=5)
buttonSBIC.place(x=1115, y=425)
#placing BO Approval Tab Wise Buttons
buttonBOA113a = tk.Button(canvas, text = 'BO-D-11(3)(a)',fg='black',bg='cyan', command = BO_approve_set)
buttonBOA113a.grid(row=5,column=5)
buttonBOA113a.place(x=20, y=490)
buttonBOA13a = tk.Button(canvas, text = 'BO-D-13(a)',fg='black',bg='cyan', command = BO_approve_rej)
buttonBOA13a.grid(row=5,column=5)
buttonBOA13a.place(x=180, y=490)
buttonBOA92 = tk.Button(canvas, text = 'BO-D-9(2)',fg='black',bg='cyan', command = BO_approve_92)
buttonBOA92.grid(row=5,column=5)
buttonBOA92.place(x=180, y=540)
buttonBOA93a = tk.Button(canvas, text = 'BO-D-9(3)(a)',fg='black',bg='cyan', command = BO_approve_93a)
buttonBOA93a.grid(row=5,column=5)
buttonBOA93a.place(x=20, y=540)
buttonBOAIA = tk.Button(canvas, text = 'BO-Iss. Adv.',fg='black',bg='cyan', command = IssueAdvisory)
buttonBOAIA.grid(row=5,column=5)
buttonBOAIA.place(x=20, y=590)
buttonBO93c = tk.Button(canvas, text = 'BO-D-9(3)(c)',fg='black',bg='cyan', command = BO_approve_93c)
buttonBO93c.grid(row=5,column=5)
buttonBO93c.place(x=180, y=590)
buttonBO35 = tk.Button(canvas, text = 'BO-D-3(5)',fg='black',bg='cyan', command = BO_approve_35)
buttonBO35.grid(row=5,column=5)
buttonBO35.place(x=180, y=450)
#placing NBFCO Approval Tab Wise Buttons
buttonNA114a = tk.Button(canvas, text = 'N-D-11(4)(a)',fg='black',bg='cyan', command = NBFCO_114a)
buttonNA114a.grid(row=5,column=5)
buttonNA114a.place(x=710, y=490)
buttonNA131a = tk.Button(canvas, text = 'N-D-13(1)(a)',fg='black',bg='cyan', command = NBFCO_131a)
buttonNA131a.grid(row=5,column=5)
buttonNA131a.place(x=850, y=490)
buttonNA132 = tk.Button(canvas, text = 'N-D-13(2)',fg='black',bg='cyan', command = NBFCO_132)
buttonNA132.grid(row=5,column=5)
buttonNA132.place(x=710, y=540)
buttonNA91a = tk.Button(canvas, text = 'N-D-9(1)(a)',fg='black',bg='cyan', command = NBFCO_91a)
buttonNA91a.grid(row=5,column=5)
buttonNA91a.place(x=850, y=540)
buttonNA9Aa = tk.Button(canvas, text = 'N-D-9(A)(a)',fg='black',bg='cyan', command = NBFCO_9Aa)
buttonNA9Aa.grid(row=5,column=5)
buttonNA9Aa.place(x=710, y=590)
buttonNOA9ac = tk.Button(canvas, text = 'N-D-9(A)(c)',fg='black',bg='cyan', command = NBFCO_9ac)
buttonNOA9ac.grid(row=5,column=5)
buttonNOA9ac.place(x=850, y=590)
buttonNAIA = tk.Button(canvas, text = 'N-Iss. Adv.',fg='black',bg='cyan', command = NBFCO_IssueAdvisory)
buttonNAIA.grid(row=5,column=5)
buttonNAIA.place(x=975, y=490)
buttonNA35 = tk.Button(canvas, text = 'N-D-3(5)',fg='black',bg='cyan', command = NBFCO_35)
buttonNA35.grid(row=5,column=5)
buttonNA35.place(x=975, y=530)
buttonNA91 = tk.Button(canvas, text = 'N-D-9(1)',fg='black',bg='cyan', command = NBFCO_91)
buttonNA91.grid(row=5,column=5)
buttonNA91.place(x=1085, y=490)
buttonNA9Ag = tk.Button(canvas, text = 'N-D-9(A)(g)',fg='black',bg='cyan', command = NBFCO_9Ag)
buttonNA9Ag.grid(row=5,column=5)
buttonNA9Ag.place(x=1075, y=530)
# Placing search buttons and various search button related operations
buttonLOADED101 = tk.Button(canvas, text = 'Already Loaded IR to 10-1',fg='black',bg='cyan', command = to101_loaded)
buttonLOADED101.grid(row=5,column=5)
buttonLOADED101.place(x=380, y=595)
entryS = ttk.Entry(canvas, width = 30)
entryS.grid(row=5,column=5)
entryS.place(x = 370,y = 490)
buttonS = ttk.Button(canvas, text = 'Search as View',style="C.TButton", command = Search)
buttonS.grid(row=5,column=5)
buttonS.place(x = 300, y = 520)
#Add edit mode search button
buttonSE = ttk.Button(canvas, text = 'Search as Edit',style="C.TButton", command = SearchE)
buttonSE.grid(row=5,column=5)
buttonSE.place(x = 500, y = 520)
#Add tick all button
buttonTA = ttk.Button(canvas, text = 'Match n Tick',style="C.TButton", command = tick_all)
buttonTA.grid(row=5,column=5)
buttonTA.place(x = 300, y = 550)
#Add LOAD Info Reqd to 10-1 button
buttonLOAD101 = ttk.Button(canvas, text = 'Load IR to 10-1',style="C.TButton", command = to101_load)
buttonLOAD101.grid(row=5,column=5)
buttonLOAD101.place(x = 500, y = 550)

canvas.create_text(540,40,text = "~CMS Manager~", font= ('Georgia', '15', 'bold'), anchor='sw')

# canvas.create_text(7,630,text= 'Note: Do not close the new chrome browser opened with this program, this is where we\'ll work. \n(All commands except closing drafts as SFA will mostly run from first tab only). Keep checking active tab by Switch to Active Tab button \n          Make sure the Secretary, DOs and NOs etc. have correct & updated data in excel files in dist folder.',font= ('Arial','11','bold','italic'), anchor='sw')

# canvas.create_text(650,540,text = 'Enter complaint no. only. To search multiple nos., seperate them by single space or just copy paste column vlaues from any excel file.',font= ('Trajanus','9', 'italic'), fill='blue')

window.mainloop()