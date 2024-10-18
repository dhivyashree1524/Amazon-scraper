import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import os
import platform  # To handle file opening across platforms

# Set up Selenium WebDriver without headless mode (so the browser is visible)
chrome_options = Options()
# Remove headless option so the browser will be shown
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

# Initialize WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Amazon India URL for laptops
amazon_url = "https://www.amazon.in/s?k=laptops"

try:
    print("Opening Amazon...")
    driver.get(amazon_url)
    time.sleep(5)  # Wait for the page to load

    # Detect if CAPTCHA is present (Amazon sometimes uses this)
    if "Enter the characters you see below" in driver.page_source:
        raise Exception("CAPTCHA detected, cannot proceed with scraping.")

    print("Page Title:", driver.title)

    # Scroll to load more products
    print("Scrolling to load more products...")
    for _ in range(3):  # Adjust the number of scrolls
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.END)
        time.sleep(3)  # Allow some time for new items to load

    # Parse the page with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    products = []

    print("Extracting product data...")
    for item in soup.find_all("div", {"data-component-type": "s-search-result"}):
        title = item.h2.get_text(strip=True) if item.h2 else "Title not found"

        # Attempt to extract price
        price = item.find('span', class_='a-price')
        if price:
            whole = price.find('span', class_='a-price-whole')
            fraction = price.find('span', class_='a-price-fraction')
            if whole and fraction:
                price_value = f"{whole.get_text(strip=True)}.{fraction.get_text(strip=True)}"
            else:
                price_value = price.get_text(strip=True)
        else:
            price_value = "Not available"

        # Attempt to extract rating
        rating = item.find('span', class_='a-icon-alt')
        rating_value = rating.get_text(strip=True) if rating else "No rating"

        # Attempt to extract number of reviews
        review_count = item.find('span', class_='a-size-base')
        reviews_value = review_count.get_text(strip=True) if review_count else "No reviews"

        # Attempt to extract availability status (usually in stock status or unavailable)
        availability = item.find('span', class_='a-declarative')
        availability_value = "Available" if availability and "In stock" in availability.get_text(strip=True) else "Unavailable"

        # Debugging information
        print(f"Title: {title}, Price: {price_value}, Rating: {rating_value}, Reviews: {reviews_value}, Availability: {availability_value}")

        # Collect product data if price is available
        if title and price_value != "Not available":
            products.append({
                'Product Title': title,
                'Price (INR)': price_value,
                'Rating': rating_value,
                'Reviews': reviews_value,
                'Availability': availability_value
            })

    if not products:
        print("No products found. Please check the Amazon page structure or the search query.")
    else:
        # Save product data to multiple file formats
        df = pd.DataFrame(products)
        df.to_csv('amazon_laptops.csv', index=False)
        df.to_json('amazon_laptops.json', orient='records', lines=True)
        df.to_excel('amazon_laptops.xlsx', index=False)

        # Close the browser automatically after scraping is done
        print("Closing the browser...")
        driver.quit()

        # Open files on appropriate platforms
        if platform.system() == 'Windows':
            os.startfile('amazon_laptops.csv')
            os.startfile('amazon_laptops.json')
            os.startfile('amazon_laptops.xlsx')
        elif platform.system() == 'Darwin':  # macOS
            os.system(f'open amazon_laptops.csv')
            os.system(f'open amazon_laptops.json')
            os.system(f'open amazon_laptops.xlsx')
        else:  # Linux or other
            os.system(f'xdg-open amazon_laptops.csv')
            os.system(f'xdg-open amazon_laptops.json')
            os.system(f'xdg-open amazon_laptops.xlsx')

        print(f"Data saved: {len(products)} products found.")

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    # Ensure the browser is closed in case of any errors
    if driver.service.process is not None:
        driver.quit()