import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import requests
import os
from datetime import datetime
import json

chrome_options = Options()
chrome_options.add_argument("--start-maximized")  # Maximize the window
chrome_options.add_argument("--disable-notifications")  # Disable browser notifications

# Set up the WebDriver (Chrome)
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Read the Excel file containing product page links
file_path = 'Box_Links.xlsx'  # Update the file path if needed
df = pd.read_excel(file_path)

# Wait function to handle elements
def wait_for_element(xpath_selector):
    try:
        WebDriverWait(driver, 30).until(  # Increased the timeout to 30 seconds
            EC.presence_of_element_located((By.XPATH, xpath_selector))  # Wait for element presence first
        )
        WebDriverWait(driver, 30).until(  # Additional wait for visibility
            EC.visibility_of_element_located((By.XPATH, xpath_selector))
        )
    except Exception as e:
        print(f"execute: {e}")

# Refresh the page to ensure no caching issues
def refresh_page():
    driver.refresh()
    time.sleep(5)

def close_popups():
    try:
        # Common XPath for pop-up close buttons (adjust the XPath based on the website)
        close_button_xpath = '//*[@id="close-popup-button"]'  # Replace with the actual XPath
        wait_for_element(close_button_xpath)
        
        # Close the pop-up by clicking the close button
        close_button = driver.find_element(By.XPATH, close_button_xpath)
        close_button.click()
        print("Pop-up closed successfully.")
    except Exception as e:
        print(f"No pop-up detected or failed to close: {e}")

# Function to process breadcrumbs
def process_breadcrumbs():
    breadcrumb_section_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/app-breadcrumbs/div/div/div'  # Breadcrumb container
    try:
        # Get all breadcrumb links inside the section
        breadcrumb_elements = driver.find_elements(By.XPATH, breadcrumb_section_xpath + '/div/a')
        
        # Ignore the first Home breadcrumb and the last one (MPN)
        sub_category = breadcrumb_elements[1].text.strip() if len(breadcrumb_elements) > 1 else None
        child_category = breadcrumb_elements[2].text.strip() if len(breadcrumb_elements) > 2 else None
        grand_child_categories = [breadcrumb_elements[i].text.strip() for i in range(3, len(breadcrumb_elements))]

        # Strip '>' symbols from sub_category and child_category
        if sub_category and '>' in sub_category:
            sub_category = sub_category.split('>')[-1].strip()
        if child_category and '>' in child_category:
            child_category = child_category.split('>')[-1].strip()

        # Filter out grand_child_categories that contain '>'
        grand_child_categories = [text.strip() for text in grand_child_categories if '>' not in text]
        
        return sub_category, child_category, grand_child_categories

    except Exception as e:
        print(f"Error retrieving breadcrumbs: {e}")
        return None, None, []

def scrape_specifications():
    specifications = {}

    try:
        # Expand the specifications section if it's not already expanded
        spec_tab_xpath = '//*[@id="accordion"]/p-accordion/div/p-accordiontab[2]'
        if wait_for_element(spec_tab_xpath, timeout=20):
            spec_tab = driver.find_element(By.XPATH, spec_tab_xpath)
            driver.execute_script("arguments[0].scrollIntoView(true);", spec_tab)
            spec_tab.click()
            time.sleep(2)

        # Wait for the specifications section to load
        spec_main_header_xpath = '//*[@id="index-1_header_action"]/span[2]'
        spec_main_div_xpath = '//*[@id="index-1_content"]/div/div'
        
        if not (wait_for_element(spec_main_header_xpath) and wait_for_element(spec_main_div_xpath)):
            return json.dumps({"Specs": "No specifications found"})

        main_header = driver.find_element(By.XPATH, spec_main_header_xpath).text.strip()

        tables = driver.find_elements(By.XPATH, spec_main_div_xpath + '/table')
        headers = driver.find_elements(By.XPATH, spec_main_div_xpath + '/p')
        
        specifications["MainHeader"] = main_header
        specifications["Specs"] = []

        # Loop over each table and extract key-value pairs from rows
        for i, table in enumerate(tables):
            title = headers[i].text.strip() if i < len(headers) else f"Table {i+1}"
            table_data = []
            
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                columns = row.find_elements(By.TAG_NAME, "td")
                if len(columns) == 2:
                    key = columns[0].text.strip()
                    value = columns[1].text.strip()
                    if key and value:
                        table_data.append({"Key": key, "Value": value})

            # Append the table data to the specifications JSON structure
            if table_data:
                specifications["Specs"].append({"Header": title, "Attributes": table_data})

        if not specifications["Specs"]:
            specifications["Specs"] = "No specifications found"

    except Exception as e:
        #logging.error(f"Error retrieving specifications: {e}")
        specifications = {"Specs": "Error retrieving specifications"}

    return json.dumps(specifications, indent=4)

def download_image(image_url, product_mpn, img_count, image_type="price"):
    try:
        img_data = requests.get(image_url).content
        img_folder = "product_images"
        
        if not os.path.exists(img_folder):
            os.makedirs(img_folder)
        
        # Create image filename based on Product MPN, image index, and fixed 'price' keyword
        img_name = f"{product_mpn}-{img_count}-{image_type}.jpg"  # Image name format: Product MPN + image_count + price + jpg
        img_path = os.path.join(img_folder, img_name)
        
        with open(img_path, 'wb') as f:
            f.write(img_data)
        
        print(f"Image downloaded: {img_name}")
        return img_name  # Return the image name without extension
    except Exception as e:
        print(f"Failed to download image: {e}")
        return None

def scrape_tags():
    tags = []
    try:
        tags_section_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[1]/div[2]'
        toast_elements = driver.find_elements(By.XPATH, tags_section_xpath + '/div/app-product-toast')
        
        for toast_element in toast_elements:
            toast_text = toast_element.find_element(By.XPATH, './div/span').text.strip()
            if toast_text:
                tags.append(toast_text)
        
        if not tags:
            tags.append("N/A")
        
        return {"Tags": tags}
    except Exception as e:
        print(f"Error retrieving tags: {e}")
        return {"Tags": "Error retrieving tags"}
# Function to scrape Key Features
def scrape_key_features():
    key_features = []
    try:
        key_features_section_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[2]/div[2]/ul'
        
        key_feature_elements = driver.find_elements(By.XPATH, key_features_section_xpath + '/li')
        
        for feature in key_feature_elements:
            feature_text = feature.text.strip()
            # Collect each key feature into the list
            if feature_text:
                key_features.append(feature_text)

        if not key_features:
            key_features = ["N/A"]  # If no key features found, add N/A
            # Return the key features as a JSON format
        return json.dumps({"Key_Feature": key_features}, indent=4)
    
    except Exception as e:
        print(f"Error retrieving key features: {e}")
        return json.dumps({"Key_Feature": "N/A"}, indent=4)

# Function to scrape FAQ section
def scrape_faqs():
    faqs = []

    try:
        faq_section_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[3]'
        if not wait_for_element(faq_section_xpath, timeout=20):
            return json.dumps({"FAQs": "No FAQs found"})

        faq_elements = driver.find_elements(By.XPATH, faq_section_xpath + '/app-faqs/p-accordion')

        for faq_element in faq_elements:
            question_xpath = './/div/p-accordiontab/div/div[1]/span'
            answer_xpath = './/div/p-accordiontab/div/div[2]/div'
            
            try:
                question = faq_element.find_element(By.XPATH, question_xpath).text.strip()
                answer = faq_element.find_element(By.XPATH, answer_xpath).text.strip()

                if question and answer:
                    faqs.append({"Question": question, "Answer": answer})

            except Exception as e:
                logging.warning(f"Error retrieving FAQ question or answer: {e}")
        
        if not faqs:
            faqs = [{"Question": "N/A", "Answer": "No FAQs found"}]

    except Exception as e:
        #logging.error(f"Error retrieving FAQs: {e}")
        faqs = [{"Question": "N/A", "Answer": "No FAQs found"}]

    return json.dumps({"FAQs": faqs}, indent=4)
# Data list to store all scraped product data
scraped_data = []

# Iterate over each product page link in the Excel file
for index, row in df.iterrows():
    product_link = row['Links']
 
    print(f"Opening product page: {product_link}")
    driver.get(product_link)
 
    product_name_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section/div/div[1]/div[2]/div[1]/h1'
    product_mpn_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section/div/div[1]/div[2]/div[1]/div[1]/span'  # MPN XPath
    product_price_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[1]/div[3]/div/div[1]/div[1]/span'  # Current Price XPath
    product_list_price_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[1]/div[3]/div/div[2]/span'  # List Price XPath
    breadcrumb_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section/app-breadcrumbs/div/div/div/div'  # Breadcrumbs XPath
    image_base_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[1]/div/app-custom-pdp-swiper/div[2]/div[2]/div/div[2]/div'  # Base XPath for images
    tags_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[1]/div[2]'  # XPath for tags
    
    wait_for_element(product_name_path)
    wait_for_element(product_mpn_path)
    wait_for_element(product_price_path)
    wait_for_element(product_list_price_path)
    wait_for_element(breadcrumb_path)
    wait_for_element(tags_xpath)  # Wait for tags to load
    
    for i in range(4):  # Loop to scrape the first 4 images
        wait_for_element(f"{image_base_xpath}[{i+1}]/img")  # Wait for image elements to load
       
    try:
        product_name = driver.find_element(By.XPATH, product_name_path).text
        print(f"Product Name: {product_name}")
        
        product_mpn = driver.find_element(By.XPATH, product_mpn_path).text
        product_mpn = product_mpn.replace("MPN:", "").strip()
        print(f"Product MPN: {product_mpn}")
        
        product_price = driver.find_element(By.XPATH, product_price_path).text
        product_price = product_price.replace(" INC VAT", "").strip()
        print(f"Product Current Price: {product_price}")
        
        product_list_price = driver.find_element(By.XPATH, product_list_price_path).text
        if "SAVE" in product_list_price:
            product_list_price = product_list_price.split(" SAVE")[0].strip()
        
        product_list_price = product_list_price.replace("was", "").strip()
        print(f"Product List Price: {product_list_price}")
        
        sub_category, child_category, grand_child_categories = process_breadcrumbs()
        
        print(f"Sub Category: {sub_category}")
        print(f"Child Category: {child_category}")
        if grand_child_categories:
            print(f"Grand Child Categories: {grand_child_categories}")
        
        # Download and save the images
        image_names = []
        for idx in range(4):  # Loop for the first 4 images
            image_xpath = f"{image_base_xpath}[{idx+1}]/img"
            image_url = driver.find_element(By.XPATH, image_xpath).get_attribute('src')
            print(f"Image URL {idx}: {image_url}")
            image_name = download_image(image_url, product_mpn, idx + 1, "price")
            image_names.append(image_name)
        
        # Scrape tags
        tags = scrape_tags()
         
        # Scrape specifications in JSON format
        specifications = scrape_specifications()
        
        # Scrape Key Features
        key_features = scrape_key_features()

        # Scrape FAQ section
        faqs = scrape_faqs()
        
        
        # Append the scraped data to the list
        scraped_data.append({
            'Product Name': product_name,
            'Product MPN': product_mpn,
            'Product Current Price': product_price,
            'Product List Price': product_list_price,
            'Sub Category': sub_category,
            'Child Category': child_category,
            'Grand Child Categories': grand_child_categories,
            'Thumbnail_Image': image_names[0],  # Thumbnail Image
            'Additional_Image_1': image_names[1],  # Additional Image 1
            'Additional_Image_2': image_names[2],  # Additional Image 2
            'Additional_Image_3': image_names[3],  # Additional Image 3
            'Tags': json.dumps(tags),  # Save tags as JSON in the "Tags" column
            'Specifications': specifications,  # Save Specifications in JSON format
            'Key_Features': key_features,  # Save Key Features as JSON
            'FAQs': faqs  # Save FAQs in JSON format
        })
 
    except Exception as e:
        print(f"Error retrieving product details for {product_link}: {e}")
 
    time.sleep(2)

# Generate a dynamic file name with the current date and time
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
output_filename = f"scraped_product_data_{current_time}.xlsx"

# Convert the list of data into a Pandas DataFrame
scraped_df = pd.DataFrame(scraped_data)

# Check if data is being collected
print(scraped_df)  # Debugging line before saving to Excel

# Save the DataFrame to an Excel file with a dynamic filename
scraped_df.to_excel(output_filename, index=False)

print(f"Data saved to {output_filename}")

# Close pop-ups before scraping
close_popups()

# Close the driver once all products have been scraped
driver.quit()
