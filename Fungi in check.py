import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

# 读取真菌名称从输入的Excel文件
df = pd.read_excel('C:\\Users\\Administrator\\Desktop\\ID.xlsx', header=None, names=['fungus_name'])

# 读取分类数据从Excel文件
tea_taxonomy = pd.read_excel('C:\\Users\\Administrator\\Desktop\\TEA.xlsx', header=None)[0].tolist()

driver = webdriver.Chrome()
driver.get('https://www.indexfungorum.org/Names/Names.asp')

valid_names = []
valid_info = []

for _, row in df.iterrows():
    fungus_name = row['fungus_name']
    
    search_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'SearchTerm')))
    search_box.clear()
    search_box.send_keys(fungus_name)
    
    submit_button = driver.find_element(By.NAME, 'submit')
    submit_button.click()
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    
    try:
        element = driver.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/table/tbody/tr/td/a[1]')
        element.click()
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        for item in tea_taxonomy:
            if item in soup.get_text():
                valid_names.append(fungus_name)
                info = driver.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/table/tbody/tr/td').text
                valid_info.append(info)
                break
    except:
        print(f"No result found for {fungus_name}")
    
    driver.back()

valid_df = pd.DataFrame({'fungus_name': valid_names, 'info': valid_info})
valid_df.to_excel('output.xlsx', index=False)

driver.quit()