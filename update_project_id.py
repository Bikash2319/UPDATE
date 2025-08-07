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
driver.get("https://refex.dev.gensomerp.com/login")
token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJiaWthc2guc2Fob29Ac2hhcmFqbWFuLmNvbSIsImxvZ2luX2lkIjoyNiwidXNlcl9pZCI6MzEsInVzZXJfdHlwZSI6Ik8mTSBURUFNIiwiZXhwIjoxNzUzODg5MjM1fQ.XU0HZeg9R3INtGMNR_-cdn-Ot7UBn98kMSFRZwAn6x4"
driver.execute_script(f"window.localStorage.setItem('token', '{token}');")
print("Login Successful")

file_path = "C:\\Automation\\UPDATE\\Project List.xlsx"
df = pd.read_excel(file_path, "lat_long")
data_list = df.to_dict(orient="records")
for item in data_list:
    driver.get("https://refex.dev.gensomerp.com/cluster-details")
    cluster_name = item.get("Cluster")
    cluster_name.strip()
    time.sleep(0.5)
    
    add_cluster = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@ngbtooltip='Add Cluster']")))
    add_cluster.click()
    
    cluster_name = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='cluster_name']")))
    cluster_name.clear()
    cluster_name.send_keys(cluster_name)
    
    save_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Save']")))
    save_button.click()
    time.sleep(0.5)

for item in data_list:
    driver.get("https://refex.dev.gensomerp.com/plant-management")
    project_name = item.get("site_name")
    project_name.strip()
    time.sleep(0.5)
    search = driver.find_element(By.XPATH, "//input[@placeholder='Search']")    
    search.send_keys(project_name)
    time.sleep(1)
    
    #Click on Edit icon
    edit_icon = driver.find_element(By.XPATH, "//*[@id='main-wrapper']/div[1]/div/app-plant-list/div/div/div[2]/div[2]/table/tbody/tr/td[10]/a[2]")
    time.sleep(0.5)
    edit_icon.click()
    
    
    cluster = wait.until(EC.presence_of_element_located((By.ID, "cluster")))
    time.sleep(0.5)
    #Fetch and update cluster
    cluster_value = item.get("Cluster")
    cluster.clear()
    time.sleep(0.5)
    cluster.send_keys(cluster_value)
    time.sleep(0.5)

    # #Fetch and update latitude
    # lat = item.get("Lat")
    # lat = round(float(lat), 4)
    # latitude = wait.until(EC.presence_of_element_located((By.ID, "latitude")))
    # time.sleep(0.5)
    # latitude.clear()
    # time.sleep(0.5)
    # latitude.send_keys(lat)
    
    # #Fetch and update longitude
    # long = item.get("Long")
    # long = round(float(long), 4)
    # longitude = wait.until(EC.presence_of_element_located((By.ID, "longitude")))
    # time.sleep(0.5)
    # longitude.clear()
    # time.sleep(0.5)
    # longitude.send_keys(long)
    
    time.sleep(1)
    #click on update button
    driver.find_element(By.XPATH, "//button[text() = ' Update ']").click()
    print(project_name)