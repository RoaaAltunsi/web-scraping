from asyncio import sleep
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from openpyxl import Workbook

''' ---------------------------- HELPER FUNCTIONS --------------------------------- '''
# Function to scroll down until reaching the buttom of the page
def scroll_down_to_bottom(driver):
    last_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")

    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2) 
        new_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
        if new_height == last_height:
            break
        last_height = new_height

# Function to return the HTML element of each section 
def find_parent_section_by_id(driver, id):
    try:
        return WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, f'.//main[@id="content"]//div//div//div[2]//*[@data-qa="{id}"]'))
        )
    except (TimeoutException, NoSuchElementException):
        return None


''' ---------------------------- MAIN DEFINITIONS --------------------------------- '''
# Create the Ecxel workbook and add a sheet
workbook = Workbook()
sheet = workbook.active
# Header row
sheet.append(["Link", "About", "Education", "Skills"])

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chrome_options)


''' --------------------------- EXTRACT PROFILES' LINKS --------------------------- '''

driver.get(f'https://www.google.com/search?hl=en&as_q=artificial+intelligence&as_epq=&as_oq=&as_eq=&as_nlo=&as_nhi=&lr=lang_en&cr=&as_qdr=all&as_sitesearch=xing.com%2Fprofile%2F&as_occt=any&as_filetype=&tbs=#ip=1')

# Define a wait to be used for waiting for the "Show more" button to appear
wait = WebDriverWait(driver, 10)

# Scroll down and click the "Show more" button until it's no longer available
profile_links_list = []
while True:
    try:
        scroll_down_to_bottom(driver)
        try:
            # Wait for the "Show more" button to be present
            show_more_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.T7sFge.sW9g3e.VknLRd")))
            # Click the "Show more" button
            show_more_button.click()
        except NoSuchElementException:
            break
    except Exception as e:
        break

# Extract all profiles' links 
profile_links = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, './/div[@class="MjjYud"]//a[@jsname="UWckNb"]'))
)
profile_links_list.extend([link.get_attribute('href') for link in profile_links])

print(f'length: {len(profile_links_list)}')



''' ----------------------- EXTRACT DATA FROM EACH LINK --------------------------- '''
i = 1
for profile_link in profile_links_list:

    driver.get(profile_link)


    # ____________________ ABOUT SECTION _________________________
    try:
        about = ""
        about_section = find_parent_section_by_id(driver, "about-me-section")
        if about_section:
            about = about_section[0].find_element(By.XPATH, './/p').text.strip()

    except Exception as e:
        print(f"ERROR extraxting ABOUT section: {str(e)}")

    # ___________________ EDUCATION SECTION ______________________
    try:
        education = []
        edu_entries = find_parent_section_by_id(driver, "education-entry")
        if edu_entries:
            for edu_entry in edu_entries:
                h4_element = WebDriverWait(edu_entry, 10).until(
                    EC.visibility_of_element_located((By.XPATH, './/h4'))
                )
                education.append(h4_element.text.strip())

    except Exception as e:
        print(f"ERROR extraxting EDUCATION section: {str(e)}")


    # ____________________ SKILLS SECTION ________________________
    try:
        skills = []
        skills_section = find_parent_section_by_id(driver, "skills-section")
        
        if skills_section:
            all_skills = WebDriverWait(skills_section[0], 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'list-tags_list-tagsContainer-e04a65b8'))
            )
            skill_entries = all_skills.find_elements(By.XPATH, './/div[@data-qa="skills-tags"]')
            skills.extend([skill.text for skill in skill_entries])

    except Exception as e:
        print(f"ERROR extraxting SKILLS section: {str(e)}")

    print(i)
    i=i+1
    sheet.append([profile_link, about, ', '.join(education), ', '.join(skills)])
    

# Save the Excel file
workbook.save("xing_AI_output.csv")

driver.quit()