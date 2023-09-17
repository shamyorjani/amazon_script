from time import sleep
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import openpyxl

# Define the path to the CRX file
extension_path = './amazoncrxextension.crx'

# Set up Chrome WebDriver with the extension
chrome_options = webdriver.ChromeOptions()
chrome_options.add_extension(extension_path)

# Create the WebDriver with ChromeOptions
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()

# Switch to the new tab
driver.switch_to.window(driver.window_handles[0])

# Define the base URL of Amazon.de
base_url = 'https://www.amazon.de/'
url = []
# Create a single Excel workbook and worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active


amazon_url = "https://www.amazon.de/-/en/gp/bestsellers/amazon-devices/530884031/ref=zg_bs_nav_amazon-devices_1"
sleep(5)
response = driver.get(amazon_url)

wait = WebDriverWait(driver, 5)
try:
    # Check for the presence of the "sp-cc-rejectall-link" element
    reject_button = wait.until(
        EC.element_to_be_clickable((By.ID, 'sp-cc-rejectall-link')))

    # Click the "Continue without accepting" link
    reject_button.click()
except Exception as e:
    pass
# Parse the HTML content of the page
sleep(5)
soup = BeautifulSoup(driver.page_source, "lxml")

try:
    # Navigate to the URL using Selenium
    driver.get(amazon_url)

    # Wait for the page to load and find the relevant elements
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        (By.CLASS_NAME, '_p13n-zg-nav-tree-all_style_zg-browse-group__88fbz')))

    # Parse the HTML content of the page using BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find all div elements with role="group" and class="_p13n-zg-nav-tree-all_style_zg-browse-group__88fbz"
    group_divs = soup.find_all(
        'div', {'class': '_p13n-zg-nav-tree-all_style_zg-browse-group__88fbz'})

    # print(len(group_divs))
    # Loop through each group div and extract names and URLs
    for index, group_div in enumerate(group_divs):
        # Find all div elements with role="treeitem" and class="_p13n-zg-nav-tree-all_style_zg-browse-item__1rdKf"
        categories = []
        if index == 0:
            item_divs = group_div.find_all('div', {
                'role': 'treeitem', 'class': '_p13n-zg-nav-tree-all_style_zg-browse-item__1rdKf'})

            # Loop through item divs to extract names and URLs
            for item_div in item_divs[:2]:
                # Find the anchor tag within the div
                anchor_tag = item_div.find('a')
                span_tag = item_div.find('span')

                if span_tag:
                    categories.append(
                        {'name': span_tag.text.strip(), 'url': ''})

                if anchor_tag:
                    category_name = anchor_tag.text.strip()  # Extract category name
                    # Append base URL to category URL
                    category_url = base_url + anchor_tag['href']
                    categories.append(
                        {'name': category_name, 'url': category_url})

            # print(categories)
            for category in categories:
                url.append({"name": category['name'], "url": category['url']})
        if index == 2:
            item_divs = group_div.find_all('div', {
                'role': 'treeitem', 'class': '_p13n-zg-nav-tree-all_style_zg-browse-item__1rdKf'})

            # Loop through item divs to extract names and URLs
            for item_div in item_divs:
                # Find the anchor tag within the div
                anchor_tag = item_div.find('a')
                span_tag = item_div.find('span')

                if span_tag:
                    categories.append(
                        {'name': span_tag.text.strip(), 'url': ''})

                if anchor_tag:
                    category_name = anchor_tag.text.strip()  # Extract category name
                    # Append base URL to category URL
                    category_url = base_url + anchor_tag['href']
                    categories.append(
                        {'name': category_name, 'url': category_url})
            # print(categories)
            for category in categories:
                url.append({"name": category['name'], "url": category['url']})

except TimeoutException:
    print("Timed out waiting for page to load, retrying...")

for ur in url:
    print(ur)
