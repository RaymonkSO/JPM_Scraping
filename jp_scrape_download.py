import os
import openpyxl
import time
from pathlib import Path
from selenium.webdriver import ActionChains
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.edge.options import Options
from bs4 import BeautifulSoup as BS

#Setting Up Relative Path
parent_dir = Path(__file__).parent
download_dir = str(parent_dir / "download")
chromedriver_path = os.path.join(parent_dir, 'chromedriver.exe')


#Clean-up File
for files in os.listdir(download_dir):
    file_path = os.path.join(download_dir, files)
    if os.path.isfile(file_path):
        os.remove(file_path)

#Set Up Excel
wbook = openpyxl.load_workbook("list_unsort.xlsx")
ws = wbook["Sheet1"]
keywords = []
for row in ws.iter_rows(values_only=True):
    cleaned_row = [str(value).strip() if value else '' for value in row]
    row_string = ' '.join(cleaned_row)
    keywords.append(row_string)


# Set Driver Options and Service
gdriver_opt = webdriver.ChromeOptions()
gdriver_opt.add_argument('--start-maximized')
gdriver_opt.add_argument('--disable-extensions')
gdriver_opt.add_argument("--lang=en-GB")
# gdriver_opt.add_argument("--headless")
# gdriver_opt.add_argument("--incognito")
gdriver_opt.add_experimental_option("detach", True)
gdriver_opt.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
    })

gdriver_service = Service(exececutable_path=chromedriver_path)

driver = webdriver.Chrome(service=gdriver_service, options=gdriver_opt)
wait = WebDriverWait(driver, 45)

for keyword in keywords:
    #Navigate through browser
    driver.get("https://www.google.com/search?q=" + keyword)
    html_content = driver.page_source
    parse_html = BS(html_content, "html.parser")
    web_link = parse_html.find_all('a', href = lambda href: href and "https://am.jpmorgan.com/" in href)
    web_link_urls = [link.get('href') for link in web_link]
    driver.get(web_link_urls[0])

    #Go to History URL
    cur_url = driver.current_url
    historyword = "#historical-price"
    history_url = cur_url + historyword
    driver.get(history_url)

    #Accept All Pop Ups
    try:
        loc = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[3]/button[2]')
        loc.click()
    except NoSuchElementException:
        pass
    
    try:
        onebts = driver.find_element(By.XPATH, '//*[@id="onetrust-accept-btn-handler"]')
        onebts.click()
    except NoSuchElementException:
        pass

    try:
        cookies = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/header/div/div[1]/div/div/div[3]/div/a[1]')
        cookies.click()
    except NoSuchElementException:
        pass

    #Find the Download Button and Download
    download = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="nav-history-excel-download"]')))
    time.sleep(1)
    html_content = driver.page_source
    soup = BS(html_content, "html.parser")

    with open('test.txt', "w", encoding ='utf-8') as file:
        file.write(html_content)

    link_element = soup.find('a', {'id': 'nav-history-excel-download'})
    link = link_element.get('href')

    with open('test.txt', "w", encoding ='utf-8') as file:
        file.write(str(link_element))

    JPM_web = "https://am.jpmorgan.com/"
    link = JPM_web + link
    driver.get(link)

driver.quit()








