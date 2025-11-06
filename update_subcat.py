import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select

options = Options()
service = Service(executable_path=r"C:\\Program Files\\Python\\Python313\\Scripts\\chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
wait = WebDriverWait(driver, 10)

#Inject token for authentication
driver.get("https://refex.gensomerp.com/login")
token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJiaWthc2guc2Fob29Ac2hhcmFqbWFuLmNvbSIsImxvZ2luX2lkIjoyNiwidXNlcl9pZCI6MzEsInVzZXJfdHlwZSI6Ik8mTSBURUFNIiwiZXhwIjoxNzU1MTkyNDM0fQ.z4GERb7qiZOfk3vtgGCtsk2A5EfIayJGBTECxLCao8k"
driver.execute_script(f"window.localStorage.setItem('token', '{token}');")
print("Login Successful")

file_path = "C:\\Automation\\UPDATE\\Project List.xlsx"
df = pd.read_excel(file_path, "sub_category")
data_list = df.to_dict(orient="records")
for item in data_list:
    driver.get("https://refex.gensomerp.com/subcategory-master")
    add_sub_cat_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@ngbtooltip='Add SubCategory']")))
    time.sleep(0.5)
    add_sub_cat_btn.click()
    
    time.sleep(1)
    category = item.get("cat_name")
    category.strip()
    cat_dd = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='catType']")))
    cat_dd.send_keys(category)
    cat_dd.click()
    
    time.sleep(1)
    sub_cat_name = item.get("sub_cat_name")
    sub_cat_name.strip()
    sub_cat = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='subCat']")))
    sub_cat.send_keys(sub_cat_name)

    save = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Save']")))
    time.sleep(1)
    save.click()
    print(f'{category} ----> {sub_cat_name}')
    
    