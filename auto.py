import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tqdm import tqdm
import time
import concurrent.futures

# Function to perform Google search and find the official website
def find_official_website(nbfc_name, cache):
    if nbfc_name in cache:
        return cache[nbfc_name]

    options = FirefoxOptions()
    options.headless = True  # Run in headless mode
    options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    service = FirefoxService(r'C:\Users\mail2\Downloads\geckodriver-v0.34.0-win32\geckodriver.exe')  # Update with the path to your geckodriver
    driver = webdriver.Firefox(service=service, options=options)

    try:
        driver.get("https://www.google.com")
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(f"{nbfc_name} official website")
        search_box.send_keys(Keys.RETURN)

        # Wait for the search results to load
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div#search")))

        # Extract the first search result URL
        first_result = driver.find_element(By.CSS_SELECTOR, "div#search a")
        official_website = first_result.get_attribute("href")

        cache[nbfc_name] = official_website
        return official_website
    except Exception as e:
        print(f"Error finding website for {nbfc_name}: {e}")
        cache[nbfc_name] = None
        return None
    finally:
        driver.quit()

# Main function to process the Excel file
def process_excel_file(input_file):
    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Add a new column for the official website
    df['Official Website'] = None

    # Cache for search results
    cache = {}

    # Iterate through each NBFC and find the official website using parallel processing
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {executor.submit(find_official_website, row['NBFC Name'], cache): index for index, row in df.iterrows()}
        for future in tqdm(concurrent.futures.as_completed(futures), total=len(futures), desc="Processing NBFCs"):
            index = futures[future]
            df.at[index, 'Official Website'] = future.result()

    # Write the updated DataFrame back to the same Excel file
    df.to_excel(input_file, index=False)

# Example usage
input_file = 'input.xlsx'  # You can change this to any file name
process_excel_file(input_file)# Measure the time taken to run the code
start_time = time.time()
process_excel_file(input_file)
end_time = time.time()

# Calculate and display the time taken
time_taken = end_time - start_time
print(f"Time taken to run the code: {time_taken:.2f} seconds")