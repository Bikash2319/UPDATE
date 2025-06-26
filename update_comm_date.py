import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select

chrome_options = Options()
service = Service(executable_path=r"C:\\Program Files\\Python313\Scripts\\chromedriver.exe")
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.maximize_window()
wait = WebDriverWait(driver, 10)

#Inject token for authentication
driver.get("https://refex.dev.gensomerp.com/login")
token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJhc2hpc2gua0BzaGFyYWptYW4uY29tIiwibG9naW5faWQiOjMsInVzZXJfaWQiOjMsInVzZXJfdHlwZSI6Ik8mTSBURUFNIiwiZXhwIjoxNzUwOTUxMzQxfQ.Wq5h_ZTlRrgvD9ZDwqQZ01eDaE2SCqSg7Z7GuZc8C04"
driver.execute_script(f"window.localStorage.setItem('token', '{token}');")
print("Login Successful")

file_path = "C:\\Users\\Bikash Chandra Sahoo\\OneDrive\\Desktop\\GenSOM Automation\\UPDATE\\Project List.xlsx"
df = pd.read_excel(file_path, "Sheet1")
data_list = df.to_dict(orient="records")

for item in data_list:
    driver.get("https://refex.dev.gensomerp.com/plant-management")
    plant_name = item.get("Project Name")
    search = driver.find_element(By.XPATH, "//input[@placeholder='Search']")
    search.send_keys(plant_name)
    time.sleep(1)

    edit_icon = driver.find_element(By.XPATH, "//*[@id='main-wrapper']/div[1]/div/app-plant-list/div/div/div[2]/div[2]/table/tbody/tr/td[10]/a[2]")
    edit_icon.click()

    time.sleep(1)
    comm_date = item.get("Commissioning Date")
    commission_date = driver.find_element(By.ID, "commissioning_date")
    commission_date.send_keys(comm_date)

    time.sleep(1)
    cluster_dd = Select(driver.find_element(By.ID, "cluster"))
    cluster_dd.select_by_visible_text("REIL")

    time.sleep(1)
    company_dd = Select(driver.find_element(By.ID, "company"))
    company_dd.select_by_visible_text("Refex")

    time.sleep(1)
    driver.find_element(By.XPATH, "//button[text() = ' Update ']").click()









