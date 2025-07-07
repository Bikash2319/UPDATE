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
token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJiaWthc2guc2Fob29Ac2hhcmFqbWFuLmNvbSIsImxvZ2luX2lkIjoyNiwidXNlcl9pZCI6MzEsInVzZXJfdHlwZSI6Ik8mTSBURUFNIiwiZXhwIjoxNzUxOTEzODU1fQ.zTVHyi1wGY9R2unr-chNDtscvaL6NExjZ-d9naZrtxg"
driver.execute_script(f"window.localStorage.setItem('token', '{token}');")
print("Login Successful")

file_path = "C:\\Automation\\UPDATE\\Project List.xlsx"
df = pd.read_excel(file_path, "Main")
data_list = df.to_dict(orient="records")

for item in data_list:
    try:
        driver.get("https://refex.gensomerp.com/plant-management")
        project_name = item.get("site_name")
        time.sleep(1)
        search = driver.find_element(By.XPATH, "//input[@placeholder='Search']")
        search.send_keys(project_name)
        time.sleep(2)

        edit_icon = driver.find_element(By.XPATH, "//*[@id='main-wrapper']/div[1]/div/app-plant-list/div/div/div[2]/div[2]/table/tbody/tr/td[10]/a[2]")
        edit_icon.click()

        time.sleep(1)
        comm_date = item.get("commission_date")
        date = comm_date.strftime("%Y-%m-%d")
        commission_date = driver.find_element(By.ID, "commissioning_date")
        commission_date.send_keys(date)

        # site_add = item.get("address")
        # site_address = driver.find_element(By.ID, "site_address")
        # site_address.clear()
        # site_address.send_keys(site_add)

        time.sleep(1)
        cluster_dd = Select(driver.find_element(By.ID, "cluster"))
        cluster_dd.select_by_visible_text("REIL")

        time.sleep(1)
        comp_name = item.get("company_name")
        comp_dd = wait.until(EC.element_to_be_clickable((By.ID, "company")))
        comp_dd.send_keys(comp_name)
        comp_dd.click()
        
        p_id = item.get("project_id")
        project_id = driver.find_element(By.ID, "plant_code")
        time.sleep(0.5)
        project_id.clear()
        time.sleep(0.5)
        project_id.send_keys(p_id)

        time.sleep(1)
        driver.find_element(By.XPATH, "//button[text() = ' Update ']").click()
        print(project_name)
    except Exception as e:
        print(f"Failed to update the '{project_name}")









