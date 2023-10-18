#!/usr/bin/env python
# coding: utf-8

# In[11]:


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd


# In[18]:


driver = webdriver.Chrome()
url = "https://shura2023.elections.om/voter/final-list"
driver.get(url)




region_dropdown = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ddlRegion"))
)

region_dropdown = Select(region_dropdown)
region_dropdown.select_by_value("1")  # 1 = محافظة المسقط


driver.implicitly_wait(2)  

wilayat_dropdown = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ddlWilayat"))
)


wilayat_dropdown = Select(wilayat_dropdown)
wilayat_dropdown.select_by_value("1") # 1 = ولاية المسقط




search_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "btnSearch"))
)
search_button.click()





Table_Excel = pd.DataFrame(columns=["الاسم", "الرقم التسلسلي"])  # Specify the desired column names

i = 0
while True:
    i = i+1
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div/div[1]/div/div[3]/div[2]/div/div[1]/table/tbody/tr"))
    )

    page_source = driver.page_source


    soup = BeautifulSoup(page_source, "html.parser")

    
    table = soup.find('table')

    
    data = []
    for row in table.find_all('tr'):
        row_data = [td.text.strip() for td in row.find_all('td')]
        if len(row_data) == 2:  
            data.append(row_data)

    
    df = pd.DataFrame(data, columns=["الاسم", "الرقم التسلسلي"]) 
    
    Table_Excel = pd.concat([Table_Excel, df], ignore_index=True)
    Table_Excel.to_excel("Data_Election-Of-Members.xlsx", index=False)
    
    next_button = driver.find_elements(By.XPATH, "/html/body/div[1]/div[1]/div/div[3]/div[2]/div/div[2]/a[3]")
    if not next_button:
        break

    next_button[0].click()
    print(f"Extract :{i} Table ")

driver.quit()


# In[ ]:




