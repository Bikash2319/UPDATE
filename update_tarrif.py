import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException

options = Options()
service = Service(executable_path=r"C:\\Program Files\\Python\\Scripts\\chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
driver.implicitly_wait(5)
wait = WebDriverWait(driver, 10)

#Inject token for authentication
driver.get("https://refex.gensomerp.com/login")
token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJiaWthc2guc2Fob29Ac2hhcmFqbWFuLmNvbSIsImxvZ2luX2lkIjoyNiwidXNlcl9pZCI6MzEsInVzZXJfdHlwZSI6Ik8mTSBURUFNIiwiZXhwIjoxNzYyNTI3MzI2fQ.3pjyvyVvL29MLfMnNppcgNkawlx7C0sMtYx5fNUUFsQ"
driver.execute_script(f"window.localStorage.setItem('token', '{token}');")
print("Login Successful")

file_path = "C:\\Automation\\UPDATE\\Project List.xlsx"
df = pd.read_excel(file_path, "Tarrif")
data_list = df.to_dict(orient="records")

for item in data_list:
    driver.get("https://refex.gensomerp.com/plant-management")

    try:
        project_name = item.get("site_name").strip()
        time.sleep(0.5)

        search = driver.find_element(By.XPATH, "//input[@placeholder='Search']")
        search.clear()
        search.send_keys(project_name)
        time.sleep(1)

        edit_icon = driver.find_element(By.XPATH,
            "//*[@id='main-wrapper']/div[1]/div/app-plant-list/div/div/div[2]/div[2]/table/tbody/tr/td[10]/a[2]"
        )
        time.sleep(0.5)
        edit_icon.click()
        time.sleep(0.5)

    
        new_tarrif = item.get("tarrif_changes")
        tarrif_field = driver.find_element(By.ID, "tariff")
        time.sleep(0.5)
        tarrif_field.clear()
        time.sleep(0.5)
        tarrif_field.send_keys(new_tarrif)
        time.sleep(0.5)

        driver.find_element(By.XPATH, "//button[text() = ' Update ']").click()

    
    except NoSuchElementException:
        print(f"Project '{project_name}' not found.")
        continue
    
    except Exception as e:
        print(f"Error on '{project_name}': {e}")
        continue