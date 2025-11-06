import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

options = Options()
service = Service(executable_path=r"C:\\Program Files\\Python\\Python313\\Scripts\\chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
wait = WebDriverWait(driver, 10)

#Inject token for authentication
driver.get("https://refex.dev.gensomerp.com/login")
token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJiaWthc2guc2Fob29Ac2hhcmFqbWFuLmNvbSIsImxvZ2luX2lkIjoyNiwidXNlcl9pZCI6MzEsInVzZXJfdHlwZSI6Ik8mTSBURUFNIiwiZXhwIjoxNzU2NDk3NDEzfQ.Rarkpjc3S8sFGdWge5NQ0KyoT4_5gQgG5E0uIlrQaFA"
driver.execute_script(f"window.localStorage.setItem('token', '{token}');")
print("Login Successful")

file_path = "C:\\Automation\\UPDATE\\Project List.xlsx"
df = pd.read_excel(file_path, "ac_dc")
data_list = df.to_dict(orient="records")
driver.get("https://refex.dev.gensomerp.com/inventory-managment")
for item in data_list:
    time.sleep(1)
    # driver.get("https://refex.dev.gensomerp.com/inventory-managment")
    
    warehouse_name = item.get("site_name")
    warehouse_dd = wait.until(EC.presence_of_element_located((By.ID, "dropdown_warehouse_id")))
    warehouse_dd.click()
    warehouse_dd.send_keys(warehouse_name)
    time.sleep(1)
    
    equip_name = item.get("device_name")
    search = driver.find_element(By.XPATH, "//input[@placeholder='Search']")
    time.sleep(0.5)
    search.clear()
    search.send_keys(equip_name)
    time.sleep(2)
    # search.send_keys(Keys.ENTER) 
    
    edit_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='main-wrapper']/div[1]/div/app-inventory-managment/div/div/div[2]/div[2]/table/tbody/tr/td[17]/a[2]")))
    time.sleep(1)
    edit_icon.click()
    
    dc = item.get("dc_capacity")
    print(dc)
    dc_cap = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='dc_capacity']")))
    time.sleep(0.5)
    dc_cap.clear()
    time.sleep(0.5)
    dc_cap.send_keys(dc)

    
    ac = item.get("ac_capacity")
    print(ac)
    ac_cap = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='ac_capacity']")))
    time.sleep(0.5)
    ac_cap.clear()
    time.sleep(0.5)
    ac_cap.send_keys(ac)
    
    # inv_type = item.get("inv_type")
    # inv_dd = Select(By.XPATH, "//select[@formcontrolname='inverter_type']")
    # time.sleep(0.5)
    # inv_dd.select_by_visible_text(inv_type)
    
    update = driver.find_element(By.XPATH, "//button[text() =' Update ']")
    time.sleep(0.5)
    update.click()
    print(equip_name)
    
    
    