from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException
from bs4 import BeautifulSoup
import time
import os
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook

# Set up the ChromeDriver service using webdriver-manager
service = Service(ChromeDriverManager().install())

# Initialize the WebDriver
driver = webdriver.Chrome(service=service)

try:
    # Navigate to the website
    driver.get('https://www.bernhardt.com')

    # Wait for the "Log In / Join" button to be clickable and click it
    login_button = WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.dropdown-toggle.btn.btn-link2.btn-xs[login-prompt="emun"]'))
    )
    login_button.click()

    # Wait for the email field to be visible
    email_field = WebDriverWait(driver, 3).until(
        EC.visibility_of_element_located((By.ID, 'email'))
    )

    # Locate the password field
    password_field = driver.find_element(By.ID, 'password')

    # Enter the username and password
    email_field.send_keys('USER')  # Replace with the actual username
    password_field.send_keys('PASS')  # Replace with the actual password

    # Locate and click the submit button
    submit_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
    submit_button.click()

    # Array of product categories to iterate through
    luxury_items = [
        "luxury-bedroom-furniture",
        "luxury-dining-room-furniture",
        "luxury-living-room-furniture",
        "luxury-home-office-room-furniture",
    ]
    categories = [
        "Bedroom",
        "Dining Room",
        "Living Room",
        "Office",
    ]
    item_category = []
    skus = []

    # Create a new Excel workbook and select the active worksheet
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Bernhardt Products'

    # Add header row to the sheet
    sheet.append(['Image', 'SKU', 'Name', 'Collection', 'Category', 'Descrition', 'About ', 'Features', 'Status', 'Available', 'Width', 'Depth', 'Height', 'Weight', 'Our Cost', 'Min Cost', 'Max Cost', 'Retail', 'MSRP', 'Gradeable', 'A', 'B', 'C', 'COM', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V'])

    # Iterate through each product category in the array
    for item in luxury_items:
        # Construct the URL
        url = f'https://www.bernhardt.com/products/{item}'

        # Navigate to the URL
        driver.get(url)
        
        while True:
            # Wait for the page to load and scrape the product information as needed
            time.sleep(2)  # Example wait time, adjust as necessary
                
            # Use BeautifulSoup to parse the page
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            
            products_sku = soup.find_all('span', class_='product-id ng-binding')

            for product in products_sku:
                skus.append(product.text)
                
                for i in range(len(luxury_items)):
                    if item == luxury_items[i]:
                        item_category.append(categories[i])

            # Find the text showing the current range of items
            range_text = soup.find('p', class_='ng-binding ng-scope').get_text()
            if range_text:
                range_numbers = range_text.split()
                if range_numbers[0] == "Showing":
                    current_last_item = int(range_numbers[3])
                    total_items = int(range_numbers[5])
                
                    # Check if the current page is the last page
                    if current_last_item >= total_items:
                        break
                    
                    # Find and click the "Next" button to go to the next page
                    try:
                        next_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[@ng-click='selectPage(page + 1, $event)']"))
                        )
                        next_button.click()
                    except Exception as e:
                        print(f"Error clicking next button for URL {url}: {e}")
                        break
            
    count = 0
    print(f"Total SKUs to process: {len(skus)}")
    for sku in skus:
        try:
            driver.get(f'https://www.bernhardt.com/shop/{sku}?position=-1')
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'regular-price')))
            
            # Parse the page with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Extract necessary information from the page
            item_name_element = soup.find('h1', class_='product-description ng-binding ng-scope')
            item_name = item_name_element.get_text(strip=True).strip() if item_name_element else 'Item name not found'

            description_element = soup.find('p', class_='one-up-long-desc ng-binding')
            description = description_element.get_text(strip=True).strip() if description_element else 'NA'

            about_element = soup.find('div', class_='text-center p-b-3')
            about = about_element.get_text(strip=True).strip()[15:] if about_element else 'NA'

            features_element = soup.find('div', class_='col-xs-12 column item')
            features_list = features_element.get_text(strip=True).strip().split(':') if features_element else 'NA'
            features = features_list[0] if features_list[0] != 'QTY' else 'NA'

            stock_status_element = soup.find('span', 'stock-status ng-binding ng-scope')
            stock_status = stock_status_element.get_text(strip=True).strip() if stock_status_element else 'NA'
            print(stock_status)

            # Check stock and amount in stock
            is_available_element = soup.findAll('div', class_='col-xs-4 col-sm-4 col-md-4 col-lg-3 col-xl-3 column item ng-scope')    
            count_avail = '0'
            for i in is_available_element:    
                if i.get_text(strip=True).strip()[:9] == 'AVAILABLE':
                    count_avail = i.get_text(strip=True).strip()[9:]
                    count_avail_list = count_avail.split()
                    if count_avail_list[0][-2:] == 'DC':
                        count_avail = count_avail_list[0][:-2]
                    else:
                        count_avail = count_avail_list[0]

            our_cost_element = soup.find('span', class_='regular-price ng-scope')
            our_cost = our_cost_element.find('span', class_='ng-binding').get_text(strip=True) if our_cost_element else 'Unknown Availability'
            our_cost = float(our_cost.replace(',', '').replace('$', ''))

            msrp_element = soup.find('div', class_='msrp ng-binding ng-scope')
            msrp = msrp_element.get_text(strip=True).replace("MSRP:", "").strip() if msrp_element else 'MSRP not found'
            msrp = float(msrp.replace(',', '').replace('$', ''))

            dimensions_element = soup.find('div', class_='dimensions ng-binding ng-scope')
            dimensions = dimensions_element.get_text(strip=True).strip() if dimensions_element else 'Dimensions not found'
            sizes = dimensions.split()
            width = float(sizes[1]) 
            depth = float(sizes[4])
            height = float(sizes[7])

            weight_element = soup.find('li', attrs={'ng-repeat': "labels in ::$ctrl.oneUp.product.tags['Total Shipping Weight']"})
            weight = weight_element.get_text(strip=True).strip() if weight_element else 'Weight not found'
            
            is_gradeable_element = soup.find('button', class_='btn btn-link ng-binding dropdown-toggle')
            is_gradeable = True if is_gradeable_element else False
            grade_prices = []
            if is_gradeable:
                try:
                    see_all_prices_button = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'See All Prices')]"))
                    )
                    driver.execute_script("arguments[0].click();", see_all_prices_button)
                    
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, "//td[contains(@ng-bind, '$ctrl.totalPrice(item)|currency')]"))
                    )
                except (TimeoutException, ElementClickInterceptedException) as e:
                    print(f"Error occurred while interacting with the 'See All Prices' button for SKU {sku}: {e}")
                    continue

                dropdown_button = driver.find_element(By.CSS_SELECTOR, "button.btn.btn-link.ng-binding.dropdown-toggle")
                driver.execute_script("arguments[0].click();", dropdown_button)
                
                grade_options = driver.find_elements(By.XPATH, "//ul[@class='dropdown-menu']/li/a")
                
                for index, grade_option in enumerate(grade_options):
                    driver.execute_script("arguments[0].click();", grade_option)
                    time.sleep(1)
                    
                    try:
                        price_elements = driver.find_elements(By.XPATH, "//td[contains(@ng-bind, '$ctrl.totalPrice(item)|currency')]")

                        for price_element in price_elements:
                            price_text = price_element.text.strip().replace(',', '').replace('$', '')
                            # Check if price_text is a valid float number
                            if price_text.replace('.', '', 1).isdigit():
                                grade_prices.append(float(price_text))
                            else:
                                print(f"Invalid price found for SKU {sku}: {price_text}. Setting to 0.")
                                grade_prices.append(0.0)  # Set invalid price to 0

                        driver.execute_script("arguments[0].click();", dropdown_button)
                    except Exception as e:
                        print(f"An unexpected error occurred with SKU {sku}: {e}")

            grade_prices = grade_prices[4:]

            # Find all image elements with ng-src or src attributes
            image_elements = soup.find_all('img', attrs={'ng-src': True})
            
            # Append data to sheet
            sheet.append([image_elements[0]['ng-src'], sku, " ".join(item_name.split()[1:]), item_name.split()[0], item_category[count], description, about, features, stock_status, count_avail, width, depth, height, weight, our_cost, min(grade_prices) if grade_prices else '', max(grade_prices) if grade_prices else '', ((our_cost*1.15)*2.2), msrp, is_gradeable] + grade_prices)
            for i in range(1, len(image_elements)):
                sheet.append([image_elements[i]['ng-src']]) if image_elements[i]['ng-src'][0] == 'h' else '' 
            count += 1

        except TimeoutException:
            print(f"TimeoutException: Failed to load data for SKU {sku}")
        except NoSuchElementException:
            print(f"NoSuchElementException: Element not found for SKU {sku}")
        except ElementClickInterceptedException:
            print(f"ElementClickInterceptedException: Element click intercepted for SKU {sku}")
        except WebDriverException as e:
            print(f"WebDriverException: {e}")
        except Exception as e:
            print(f"An unexpected error occurred with SKU {sku}: {e}")

finally:
    # Define the file path and save the workbook to the specified directory
    directory_path = os.path.expanduser('~/OneDrive/Documents/Coding')
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
    new_workbook_path = os.path.join(directory_path, "NewBernhardtProductInformations.xlsx")
    workbook.save(new_workbook_path)

    # Close the WebDriver
    driver.quit()
