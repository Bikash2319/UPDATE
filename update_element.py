import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys



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
df = pd.read_excel(file_path, "element")
data_list = df.to_dict(orient="records")

for item in data_list:
    
    time.sleep(1)
    driver.get("https://refex.gensomerp.com/element-master")
    add_elem = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@ngbtooltip='Add Element']")))
    time.sleep(0.5)
    add_elem.click()
    
    category = item.get("category")
    category = category.strip()
    cat_dd = driver.find_element(By.XPATH, "//select[@formcontrolname='category_id']")
    time.sleep(0.5)
    cat_dd.send_keys(category)
    cat_dd.send_keys(Keys.ENTER)
    # cat_dd = Select(driver.find_element(By.XPATH, "//select[@formcontrolname='category_id']"))
    time.sleep(0.5)

    
    sub_category = item.get("sub_category")
    sub_category = sub_category.strip()
    # sub_cat_dd = Select(driver.find_element(By.XPATH, "//select[@formcontrolname='sub_category_id']"))
    # time.sleep(1)
    # sub_opt = sub_cat_dd.select_by_visible_text(sub_category)
    sub_cat_dd = driver.find_element(By.XPATH, "//select[@formcontrolname='sub_category_id']")
    time.sleep(0.5)
    sub_cat_dd.send_keys(sub_category)
    sub_cat_dd.send_keys(Keys.ENTER)
    time.sleep(0.5)
    
    
    elem = item.get("element")
    elem = elem.strip()
    element = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='elements_name']")))
    time.sleep(0.5)
    element.send_keys(elem)
    
    
    save = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Save']")))
    time.sleep(0.5)
    save.click()
    time.sleep(1)

    
    