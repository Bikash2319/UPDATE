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
token = ""
driver.execute_script(f"window.localStorage.setItem('token', '{token}');")
print("Login Successful")

file_path = "C:\\Users\\SHARAJMAN\\Desktop\\Automation\\UPDATE\\Project List.xlsx"
df = pd.read_excel(file_path, "Main")
data_list = df.to_dict(orient="records")

for item in data_list:
    driver.get("https://refex.dev.gensomerp.com/plant-management")
    project_name = item.get("site")
    search = driver.find_element(By.XPATH, "//input[@placeholder='Search']")
    search.send_keys(project_name)
    time.sleep(1)