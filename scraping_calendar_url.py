# import chromedriver_autoinstaller
import time
import os
import pandas as pd
import numpy as np

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

from openpyxl import load_workbook
from collections import defaultdict
from bs4 import BeautifulSoup

data_folder = "Data"
if not os.path.exists(data_folder):
    os.makedirs(data_folder)

sheet_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

def select_checkbox_or_radio_by_id(element_id):
    element = driver.find_element(By.ID, element_id)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    
    if not element.is_selected():
        element.click()

def get_moon_phase(group):
    if "Full Moon" in group["Data"].values:
        group["Moon Phases"] = 1
    elif "First Qtr" in group["Data"].values or "Last Qtr" in group["Data"].values:
        group["Moon Phases"] = 2
    elif "New Moon" in group["Data"].values:
        group["Moon Phases"] = 3
    else:
        group["Moon Phases"] = 0
    return group

# chromedriver_autoinstaller.install()

url = "http://sunrisesunset.com/USA/"

# extension_path_1 = "extension_3_4_0_0.crx"
# extension_path_2 = "extension_5_10_0_0.crx"

options = Options()
# options.add_extension(extension_path_1)
# options.add_extension(extension_path_2)
options.add_argument("--headless")
options.add_argument("window-size=1920x1080")

driver = webdriver.Chrome(options=options)
# driver.maximize_window()
driver.get(url)

time.sleep(3)

main_tab = None
for tab in driver.window_handles:
    driver.switch_to.window(tab)
    if driver.current_url == url:
        main_tab = tab
        break

if main_tab is None:
    raise Exception("Main tab not found!")

for tab in driver.window_handles:
    if tab != main_tab:
        driver.switch_to.window(tab)
        driver.close()

driver.switch_to.window(main_tab)

links = driver.find_elements(By.XPATH, "//a[starts-with(@href, '/USA/') and not(contains(@href, 'print'))]")

start_navigating = False

calendar_urls = []
errors = []

for i in range(len(links)):
    links = driver.find_elements(By.XPATH, "//a[starts-with(@href, '/USA/') and not(contains(@href, 'print'))]")
    link = links[i]
    
    if "Nebraska" in link.get_attribute("href"):
        start_navigating = True

    if start_navigating:
        href = link.get_attribute("href")
        print("Visiting link:", href)
        state_name = link.text
        driver.execute_script("arguments[0].scrollIntoView();", link)
        link.click()

        try:
            dropdown_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "comb_city_info")))
            dropdown = Select(dropdown_element)

            for n, option in enumerate(dropdown.options):
                if option.text == "Choose a location":
                    continue
                
                try:
                    print(f"Selecting location: {option.text}")
                    city_name = option.text
                    dropdown.select_by_visible_text(option.text)
                    
                    checkbox_and_radio_ids = [
                        "want_twi_civ", 
                        "want_twi_naut", 
                        "want_twi_astro", 
                        "want_info", 
                        "want_mphase", 
                        "want_mrms", 
                        "want_solar_noon", 
                        "want_eqx_sol", 
                        "time_24hr", 
                        "wsos_yes"
                    ]

                    year_input = driver.find_element(By.ID, "year")
                    month_dropdown = Select(driver.find_element(By.NAME, "month"))
                    
                    year_input.clear()
                    year_input.send_keys("2022")
                    month_dropdown.select_by_value("1")
                            
                    for element_id in checkbox_and_radio_ids:
                        select_checkbox_or_radio_by_id(element_id)

                    form_element = driver.find_element(By.TAG_NAME, "form")
                    driver.execute_script("arguments[0].setAttribute('target', '_blank');", form_element)
                    submit_button = driver.find_element(By.XPATH, '//input[@type="submit" and @value="Make Calendar"]')
                    submit_button.click()

                    driver.switch_to.window(driver.window_handles[-1])
                    calendar_urls.append(driver.current_url)

                    print(driver.current_url)

                    driver.close()
                    driver.switch_to.window(main_tab)
                
                except Exception as e:
                    print(f"Error encountered for {city_name}, {state_name}: {e}")
                    errors.append((state_name, city_name))

        except NoSuchElementException:
            print(f"Dropdown not found on {href}")
        except Exception as e:
            print(f"Error encountered: {e}")

        driver.back()

driver.quit()

with open("calendar_urls.txt", "w") as file:
    for url in calendar_urls:
        file.write("%s\n" % url)

with open("errors.txt", "w") as file:
    for error in errors:
        file.write("%s, %s\n" % (error[0], error[1]))