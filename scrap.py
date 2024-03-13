from selenium import webdriver
from decouple import config
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import datetime, timedelta
import re

# Load credentials from a .env file
username = config('G_USERNAME')
password = config('G_PASSWORD')
# Initialize the Chrome browser
driver = webdriver.Chrome()

# Open the login page
driver.get("https://global-marches.com")

# Find the username and password fields and enter the credentials
username_input = driver.find_element("id", "LOGIN_INPUT")
password_input = driver.find_element("id", "PASSWORD_INPUT")

username_input.send_keys(username)
password_input.send_keys(password)

# Submit the form
password_input.send_keys(Keys.RETURN)
# Navigate to https://global-marches.com/myarchives
driver.get("https://global-marches.com/myarchives")
# Get today's date
today_date = datetime.now()

# If today is Saturday or Sunday, set DATE_LIMIT_1 to the next Monday
if today_date.weekday() == 5:  # Saturday
    date_limit_1 = today_date + timedelta(days=2)
elif today_date.weekday() == 6:  # Sunday
    date_limit_1 = today_date + timedelta(days=1)
else:
    date_limit_1 = today_date

# Convert the date to dd/mm/yyyy format
date_limit_1_str = date_limit_1.strftime("%d/%m/%Y")

# Prompt the user to input the date for DATE_LIMIT_1
user_date_1 = input("Enter the date for DATE_LIMIT_1 (dd/mm/yyyy) or press Enter for today's date: ")

# If user provides a date, use that, else use the calculated date
if user_date_1:
    date_input_1_str = user_date_1
else:
    date_input_1_str = date_limit_1_str

# Find the date picker input field for class DATE_LIMIT_1 and set the date
date_input_1 = driver.find_element("id", "DATE_LIMIT_1")
date_input_1.clear()
date_input_1.send_keys(date_input_1_str)

# Find the date picker input field for class DATE_LIMIT_2 and set today's date
date_input_2 = driver.find_element("id", "DATE_LIMIT_2")
date_input_2.clear()
date_input_2.send_keys(date_limit_1_str)

# Display available options
print("Available options:")
print("1. T11 Travaux de construction et d’aménagement")
print("2. T12 Travaux de Terrassements")
print("3. T27 Travaux de voiries, chemins, pistes et voies ferrées")

# Prompt the user to choose options
user_choices = input("Enter the options you need separated by commas (e.g., 1,2): ")

# Split the user input into a list of individual choices
choices = user_choices.split(',')

# Map user choices to corresponding option IDs
option_id_map = {
    "1": "49-selectable",
    "2": "50-selectable",
    "3": "1575-selectable",
}
for choice in choices:
    option_id = option_id_map.get(choice.strip())
    if option_id:
        # Find the option based on user input
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, option_id))
        )

        # Click on the option
        option.click()
    else:
        print(f"Invalid option: {choice.strip()}. Skipping...")
search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "SAVE"))
)
# Click the search button
search_button.click()

# Wait for the page to load after clicking the search button
time.sleep(10)  # Adjust the wait time as needed
# Wait for the class="chosen-single" element to be clickable
chosen_single_element = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CLASS_NAME, "chosen-single"))
)

# Click on the class="chosen-single" element to open the dropdown
chosen_single_element.click()

# Find and click the option with the value "500"
option_500 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//li[@class='active-result' and text()='500']"))
)
option_500.click()

# Wait for the page to load after selecting the option
time.sleep(2)  # Adjust wait time as needed
# Define nb_result_element after the search results page loads
nb_result_element = WebDriverWait(driver, 3).until(
    EC.presence_of_element_located((By.ID, "NB_RESULT"))
)

# Extract the text from nb_result_element and parse the number of results
nb_result_text = nb_result_element.text
nb_result_match = re.search(r'(\d+) Appel\(s\) d\'offre trouvé\(s\)', nb_result_text)
if nb_result_match:
    nb_result = int(nb_result_match.group(1))
else:
    raise ValueError("Unable to parse the number of results.")
# Define final_df before the loop
final_df = pd.DataFrame()
# Check if there are more than 100 results
if nb_result > 100:
    # Navigate through additional pages and click on "Suivante >"
    while True:
        # Extract data from the current page
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        result_table = soup.find('table', id='resultTable')
        data = []

        # Extract data from each row in the table
        for row in result_table.find_all('tr')[1:]:  # skipping the header row
            cells = row.find_all('td')
            num_ref = cells[1].find('span', class_='NumRef').text.strip() if cells[1].find('span', class_='NumRef') else ''
            if num_ref:  # Check if num_ref is not empty
                organism = cells[2].find('b', string='Organisme :').next_sibling.strip() if cells[2].find('b', string='Organisme :') else ''
                object_text = cells[2].find('b', string='Objet :').next_sibling.strip() if cells[2].find('b', string='Objet :') else ''
                price = cells[3].find('span', class_='c_b').text.strip() if cells[3].find('span', class_='c_b') else ''
                location = cells[4].text.strip()
                deadline = cells[5].text.strip()
                download_link = cells[7].find('a', class_='downoaldcpsicone')['href'] if cells[7].find('a', class_='downoaldcpsicone') else ''
                submission_link = cells[7].find('a', class_='btn mini btn-info')['href'] if cells[7].find('a', class_='btn mini btn-info') else ''
            if submission_link:
                submission_link = submission_link.replace("EntrepriseDetailsConsultation", "SuiviConsultation")
                # Append data to the list only if num_ref is not empty
                data.append({
                    'Num Ref': num_ref,
                    'Organism': organism,
                    'Object': object_text,
                    'Price': price,
                    'Location': location,
                    'Deadline': deadline,
                    'Download Link': download_link,
                    'Submission Link': submission_link
                })

        # Create a DataFrame from the list of data
        df = pd.DataFrame(data)

        # Append data to the final DataFrame or save it to a file
        if 'final_df' in locals():
            final_df = pd.concat([final_df, df], ignore_index=True)
        else:
            final_df = df

        # Check if there's a next page
        try:
            suivante_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Suivante >"))
            )
            suivante_button.click()
            time.sleep(5)  # Adjust wait time as needed
        except Exception as e:
            print("Reached last page or encountered an error:", e)
            break

# Save the final DataFrame to an Excel file
final_df.to_excel("global_data.xlsx", index=False)

# Close the browser window
driver.quit()
