''' ----------------------------!!! Important Note !!!----------------------------- '''
# This code allows you to only scrap 70 entry per day for the same LinkedIn account.
# If you break the rule, you will be BLOCKED by LinkedIn.
''' ------------------------------------------------------------------------------- '''

from asyncio import sleep
import random
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from parsel import Selector
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import openpyxl
from openpyxl import Workbook
 
''' ---------------------------- HELPER FUNCTIONS --------------------------------- '''
# Function to return the HTML element of each section 
def find_parent_section_by_id(driver, id):
    try:
        # Explicitly wait for the element with a specific id to be present
        div_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, id))
        )
        # Find the parent section element of about_div
        return div_element.find_element(By.XPATH, "./ancestor::section")
    except (NoSuchElementException, TimeoutException):
        return None

# Function to check if there exist more than 2 skills in skills section
def skills_show_more_existence(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "pvs-list__footer-wrapper"))
        )
        return True
    except:
        return False

''' ---------------------------- MAIN DEFINITIONS --------------------------------- '''
# Create the Ecxel workbook and add a sheet
workbook = Workbook()
sheet = workbook.active
# Header row
sheet.append(["About", "Education", "Skills"])

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
#chrome_options.add_argument("--headless")
driver = webdriver.Chrome(options=chrome_options)

''' -------------------------- LOG IN TO LINKEDIN --------------------------------- '''
driver.get("https://www.linkedin.com")
username = driver.find_element(By.ID, value="session_key")
username.send_keys('Enter your account email here') # Change it to a real value

sleep(0.5)
password = driver.find_element(By.ID, value="session_password")
password.send_keys('Enter your password here') # Change it to a real value

sleep(0.5)
sign_in_button = driver.find_element(By.XPATH, value='//*[@id="main-content"]/section[1]/div/div/form/div[2]/button')
sign_in_button.click()
sleep(15)

''' --------------------------- EXTRACT PROFILES' LINKS --------------------------- '''
# https://www.linkedin.com/search/results/people/?geoUrn=%5B%22100459316%22%2C%22101336206%22%2C%22100955028%22%2C%2290010390%22%5D&industry=%5B%22109%22%2C%221594%22%2C%2296%22%2C%226%22%2C%224%22%2C%22118%22%5D&network=%5B%22F%22%2C%22S%22%5D&origin=FACETED_SEARCH&profileLanguage=%5B%22en%22%5D&sid=*P%3B
# https://www.linkedin.com/search/results/people/?geoUrn=%5B%22100459316%22%2C%22101336206%22%2C%22100955028%22%2C%2290010390%22%5D&industry=%5B%22109%22%2C%221594%22%2C%2296%22%2C%226%22%2C%224%22%2C%22118%22%5D&network=%5B%22F%22%2C%22S%22%5D&origin=FACETED_SEARCH&profileLanguage=%5B%22en%22%5D&page={str(page)}&sid=*P%3B

page = 8
profile_links_list = []
driver.get(f'https://www.linkedin.com/search/results/people/?geoUrn=%5B%22101336206%22%5D&industry=%5B%221594%22%2C%226%22%2C%2296%22%2C%224%22%2C%22118%22%2C%22109%22%2C%223130%22%2C%225%22%5D&network=%5B%22F%22%2C%22S%22%5D&origin=FACETED_SEARCH&page=8&profileLanguage=%5B%22en%22%5D&sid=nCI')

while page != 15:
    page = page + 1

    # Wait for the elements to be present on the page
    profile_links = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'span.entity-result__title-text.t-16 a.app-aware-link'))
    )
    # Extract the href attribute for each profile and store it in a list
    profile_links_list.extend([link.get_attribute('href') for link in profile_links])

    driver.get(f'https://www.linkedin.com/search/results/people/?geoUrn=%5B%22101336206%22%5D&industry=%5B%221594%22%2C%226%22%2C%2296%22%2C%224%22%2C%22118%22%2C%22109%22%2C%223130%22%2C%225%22%5D&network=%5B%22F%22%2C%22S%22%5D&origin=FACETED_SEARCH&page={str(page)}&profileLanguage=%5B%22en%22%5D&sid=nCI')

print(f'Length: {len(profile_links_list)}')

''' ----------------------- EXTRACT DATA FROM EACH LINK --------------------------- '''

for profile_link in profile_links_list:

    driver.get(profile_link)

    # ____________________ ABOUT SECTION _________________________
    try:
        about = ""
        # Search for the parent section
        parent_section = find_parent_section_by_id(driver, "about")

        if parent_section:
            # Find the element that contains the about description by class name
            # div_element = WebDriverWait(parent_section, 10).until(
            #     EC.presence_of_element_located((By.CLASS_NAME, 'CyJJpypfBPGuJkHyhwrIsyUsAhkwOreacc'))
            # )

            # Find the span text inside the div with class=CyJJpypfBPGuJkHyhwrIsyUsAhkwOreacc
            try:
                about = parent_section.find_element(By.XPATH, './/div[@class="display-flex ph5 pv3"]//span').text.strip()
            except TimeoutException:
                about = ""

    except Exception as e:
        print(f"ERROR extraxting ABOUT section: {str(e)}")

        
        # ___________________ EDUCATION SECTION ______________________
    try:
        education = []
        # Search for the parent section
        parent_section = find_parent_section_by_id(driver, "education")
        if parent_section:

            # Find the element that contains the education level by class name
            div_elements = WebDriverWait(parent_section, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'artdeco-list__item'))
            )

            # Find the span text inside the each div with class=CRPSeLaULBpqpeYEIkQKFsqbsvacQTqJpLpyU
            for div_element in div_elements:
                try:
                    education_level = div_element.find_element(By.XPATH, ".//span[@class='t-14 t-normal']//span")
                    education.append(education_level.text.strip())
                except NoSuchElementException:
                    pass

    except Exception as e:
        print(f"ERROR extraxting EDUCATION section: {str(e)}")

        # ____________________ SKILLS SECTION ________________________
    try:
        skills=[]
        # Search for the parent section
        parent_section = find_parent_section_by_id(driver, "skills")

        if parent_section:
            # Check if there are more than 2 skills
            show_more_exist = skills_show_more_existence(parent_section)
            if show_more_exist:
                a_element = WebDriverWait(parent_section, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.pvs-list__footer-wrapper a'))
                )
                skills_link = a_element.get_attribute("href")

                # Go to skills page
                driver.get(skills_link)

                # Scroll down to trigger the loading of additional elements
                last_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")

                while True:
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)  # Adjust the sleep duration based on your needs
                    new_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
                    
                    if new_height == last_height:
                        break
                    last_height = new_height
                
                # Find the element that contains the skills by class name
                div_elements = WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, 'artdeco-list__item'))
                )
            else:
                div_elements = WebDriverWait(parent_section, 10).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, 'artdeco-list__item'))
                )

            # Find the span text inside the each div with class=CRPSeLaULBpqpeYEIkQKFsqbsvacQTqJpLpyU
            for div_element in div_elements:
                skill_text = div_element.find_element(By.XPATH, ".//span").text.strip()
                if skill_text:  # Check if the skill_text is not an empty string
                    skills.append(skill_text)

    except Exception as e:
        print(f"ERROR extraxting SKILLS section: {str(e)}")

    sheet.append([about, ', '.join(education), ', '.join(skills)])
    time.sleep(20)
    

# Save the Excel file
workbook.save("LinkedIn_Riyadh2.csv")

driver.quit()

