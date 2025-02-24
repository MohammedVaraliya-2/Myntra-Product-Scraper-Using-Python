from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import platform
import os
import shutil
import time

print(platform.architecture())

l = []

# Create directory for ChromeDriver
driver_dir = os.path.join(os.getcwd(), "chromedriver_folder")
os.makedirs(driver_dir, exist_ok=True)

chromedriver_path = os.path.join(driver_dir, "chromedriver.exe")

# Check if driver exists and is executable
if os.path.exists(chromedriver_path) and os.access(chromedriver_path, os.X_OK):
    print("✅ Using existing ChromeDriver")
else:
    print("⏳ Installing ChromeDriver...")
    chrome_install = ChromeDriverManager().install()
    shutil.move(chrome_install, chromedriver_path)
    print("✅ ChromeDriver installed in:", chromedriver_path)

# Verify driver file
if not os.path.exists(chromedriver_path) or not os.access(chromedriver_path, os.X_OK):
    raise Exception("❌ ChromeDriver installation failed or is not executable!")

# Setup WebDriver
service = ChromeService(chromedriver_path)
options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")  # Disable automation detection
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")  # Set a user-agent

try:
    driver = webdriver.Chrome(service=service, options=options)
except Exception as e:
    print("❌ Error launching ChromeDriver:", e)
    exit()

driver.get("https://www.myntra.com/mens-top")

# Wait for the page to load
wait = WebDriverWait(driver, 20)
ul_tag = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'results-base')))

html_content = driver.page_source
soup = BeautifulSoup(html_content, 'html.parser')

# Find the <ul> tag containing all the product information
ul_tag = soup.find('ul', class_='results-base')

# Find all <li> tags within the <ul> tag
product_list_items = ul_tag.find_all('li', class_='product-base')

# Iterate over each <li> tag to scrape product information
for li_tag in product_list_items:
    o = {}  # Create a new dictionary for each product
    try:
        o["Product Brand"] = li_tag.find('h3', class_='product-brand').text.strip()
    except:
        o["Product Brand"] = None
    try:
        o["Product Name"] = li_tag.find('h4', class_='product-product').text.strip()
    except:
        o["Product Name"] = None
    try:
        # Extract product price text
        price_text = li_tag.find('div', class_='product-price').find('span', class_='product-discountedPrice').text.strip()
        # Convert product price to numerical value
        o["Product Price"] = float(price_text.replace('Rs.', '').replace(',', ''))
    except:
        o["Product Price"] = None
    try:
        o["Product Ratings"] = float(li_tag.find('div', class_='product-ratingsContainer').find('span').text.strip())
    except:
        o["Product Ratings"] = None
    try:
        # Extract raw ratings count string
        ratings_count_raw = li_tag.find('div', class_='product-ratingsCount').text.strip()
        # Process the ratings count string
        if '|' in ratings_count_raw:
            ratings_count = ratings_count_raw.split("|")[1].strip()
        else:
            ratings_count = ratings_count_raw
        # Convert number of people who have rated the product to numerical value
        if 'k' in ratings_count:
            o["Number of People have rated this product"] = float(ratings_count.replace('k', '')) * 1000
        else:
            o["Number of People have rated this product"] = float(ratings_count)
    except:
        o["Number of People have rated this product"] = None

    l.append(o)

# Print the list of product info
i = 1
for product in l:
    print(i, product)
    i+=1

# Save data to Excel file
wb = Workbook()
ws = wb.active

# Write headers
headers = ["Product Brand", "Product Name", "Product Price", "Product Ratings", "Number of People have rated this product"]
ws.append(headers)

# Write data
for product in l:
    ws.append([product.get(header, "") for header in headers])

# Save workbook
wb.save("productid.xlsx")
print("Data saved to productid.xlsx")

# Close the browser
driver.quit()
