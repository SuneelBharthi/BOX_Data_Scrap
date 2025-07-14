import pandas as pd 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import traceback
import time
import requests
import os
from datetime import datetime
import json
import logging

def handle_cookie_popup(driver):
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="onetrust-accept-btn-handler"]'))
        ).click()
        print("[✓] Cookie popup handled.")
    except:
        print("[-] Cookie popup not found or already handled.")
# 2. Poster Popup (Improved version with wait)
def handle_poster_popup(driver):
    try:
        script = """
        const popup = document.querySelector("#mcforms-92356-113983");
        if (popup && popup.shadowRoot) {
            const closeBtn = popup.shadowRoot.querySelector("#el_bYfcVA1AUwL");
            if (closeBtn) {
                closeBtn.click();
                return "✅ Closed newsletter popup.";
            } else {
                return "❌ Close button not found inside shadow DOM.";
            }
        } else {
            return "❌ Popup or shadowRoot not found.";
        }
        """
        result = driver.execute_script(script)
        print(result)
        time.sleep(1)
    except Exception as e:
        print(f"❌ Exception while handling newsletter popup: {e}")

# Logging setup
log_file = f"scraping_log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger()
MAX_RETRIES = 1
def get_chrome_options():
    options = Options()
    options.add_argument("--headless=new")  # ✅ Stable headless mode
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-dev-shm-usage")
    prefs = {"profile.managed_default_content_settings.images": 2}
    options.add_experimental_option("prefs", prefs)
    #options.page_load_strategy = 'eager'
    options.add_argument("--log-level=3")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    return options

# Read the Excel file containing product page links
file_path = 'Box_Links.xlsx'  
df = pd.read_excel(file_path)

# Wait function to handle elements
def wait_for_element(driver, xpath_selector):
    try:
        WebDriverWait(driver, 22).until(  # Wait for element presence
            EC.presence_of_element_located((By.XPATH, xpath_selector))
        )
        WebDriverWait(driver, 22).until(  # Additional wait for visibility
            EC.visibility_of_element_located((By.XPATH, xpath_selector))
        )
    except Exception as e:
        logger.error(f"Exception while waiting for element [{xpath_selector}]: {e}")
        logger.error(traceback.format_exc())

def scroll_to_bottom(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")    # Scroll down to load all content on the page
    time.sleep(2) 
    driver.execute_script("window.scrollTo(0, 0);")                             # Scroll up to the top of the page
    time.sleep(2) 

# Main scraping function 
def scrape_product_page(product_link):
    options = get_chrome_options()
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 20)
    try:
        driver.get(product_link)
        scroll_to_bottom(driver)
        logger.info(f"Scraping {product_link}")
        handle_cookie_popup(driver)
        handle_poster_popup(driver) # Dismiss newsletter popup

        product_name_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section/div/div[1]/div[2]/div[1]/h1'
        product_mpn_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section/div/div[1]/div[2]/div[1]/div[1]/span'
        product_price_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[1]/div[3]/div/div[1]/div[1]/span'
        product_list_price_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[1]/div[3]/div/div[2]/span'
        breadcrumb_path = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section/app-breadcrumbs/div/div/div/div'
        image_base_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[1]/div/app-custom-pdp-swiper/div[2]/div[2]/div/div[2]/div'
        tags_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/div/div[1]/div[2]/div[1]/div[2]'

        wait_for_element(driver, product_name_path)
        wait_for_element(driver, product_mpn_path)
        wait_for_element(driver, product_price_path)
        wait_for_element(driver, product_list_price_path)
        wait_for_element(driver, breadcrumb_path)
        wait_for_element(driver, tags_xpath)
        for i in range(4):
            wait_for_element(driver, f"{image_base_xpath}[{i+1}]/img")
        # Extract core product info
        product_name = driver.find_element(By.XPATH, product_name_path).text
        product_mpn = driver.find_element(By.XPATH, product_mpn_path).text.replace("MPN:", "").strip()
        product_price = driver.find_element(By.XPATH, product_price_path).text.replace(" INC VAT", "").strip()
        product_list_price = driver.find_element(By.XPATH, product_list_price_path).text
        if "SAVE" in product_list_price:
            product_list_price = product_list_price.split(" SAVE")[0].strip()
        product_list_price = product_list_price.replace("was", "").strip()
        # Breadcrumbs
        category, sub_category, child_categories = process_breadcrumbs(driver)

        image_names = []
        for idx in range(4):
            try:
                image_xpath = f"{image_base_xpath}[{idx+1}]/img"
                image_url = driver.find_element(By.XPATH, image_xpath).get_attribute('src')
                image_name = download_image(image_url, product_mpn, None if idx == 0 else idx, "price")
                image_names.append(image_name)
            except Exception as e:
                logger.warning(f"⚠️ Failed to get image at index {idx+1}: {e}")
                image_names.append(None)
        # Scrape additional details
        tags = scrape_tags(driver)
        key_features = scrape_key_features(driver)
        specifications = scrape_specifications(driver, wait)
        faqs = scrape_faqs(driver)
        return {
            "Link": product_link,
            'Product Name': product_name,
            'Product MPN': product_mpn,
            'Product Current Price': product_price,
            'Product List Price': product_list_price,
            'Category': category,
            'Sub Category': sub_category,
            'Child Categories': child_categories,
            'Thumbnail_Image': image_names[0],
            'Additional_Image_1': image_names[1],
            'Additional_Image_2': image_names[2],
            'Additional_Image_3': image_names[3],
            'Tags': json.dumps(tags),
            'Key_Features': key_features,
            'Specifications': specifications,
            'FAQs': faqs
        }
    except Exception as e:
        logger.error(f"❌ Error scraping {product_link}: {e}")
        logger.error(traceback.format_exc())
        return None
    finally:
        driver.quit()

# Function to process breadcrumbs
def process_breadcrumbs(driver):
    breadcrumb_section_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[1]/app-breadcrumbs/div/div/div'  # Breadcrumb container
    try:
        breadcrumb_elements = driver.find_elements(By.XPATH, breadcrumb_section_xpath + '/div/a')
        category = breadcrumb_elements[1].text.strip() if len(breadcrumb_elements) > 1 else None
        sub_category = breadcrumb_elements[2].text.strip() if len(breadcrumb_elements) > 2 else None
        child_categories = [breadcrumb_elements[i].text.strip() for i in range(3, len(breadcrumb_elements))]
        if category and '>' in category:
            category = category.split('>')[-1].strip()
        if sub_category and '>' in sub_category:
            sub_category = sub_category.split('>')[-1].strip()
        child_categories = [text.strip() for text in child_categories if '>' not in text]
        return category, sub_category, child_categories
    except Exception as e:
        logger.error(f"Error retrieving breadcrumbs: {e}")
        return None, None, []

# == specifications ==
def scrape_specifications(driver, wait):
    specifications = {}
    try:
        spec_tab_xpath = '//*[@id="accordion"]/p-accordion/div/p-accordiontab[2]'
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, spec_tab_xpath)))
            spec_tab = driver.find_element(By.XPATH, spec_tab_xpath)
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", spec_tab)
            time.sleep(2)
            if not spec_tab.get_attribute("aria-expanded") == "true":
                spec_tab.click()
                time.sleep(2)
        except Exception as e:
            logger.error(f"❌ Error clicking spec tab: {e}")
            specifications["Specs"] = "Specs tab not clickable"
            return json.dumps(specifications, indent=4)
        spec_main_header_xpath = '//*[@id="index-1_header_action"]/span[2]'
        spec_main_div_xpath = '//*[@id="index-1_content"]/div/div'
        wait.until(EC.presence_of_element_located((By.XPATH, spec_main_header_xpath)))
        wait.until(EC.presence_of_element_located((By.XPATH, spec_main_div_xpath)))
        main_header = driver.find_element(By.XPATH, spec_main_header_xpath).text.strip()
        specifications["MainHeader"] = main_header
        specifications["Specs"] = []
        tables = driver.find_elements(By.XPATH, spec_main_div_xpath + '/table')
        headers = driver.find_elements(By.XPATH, spec_main_div_xpath + '/p')
        for i, table in enumerate(tables):
            title = headers[i].text.strip() if i < len(headers) else f"Table {i+1}"
            table_data = []
            for row in table.find_elements(By.TAG_NAME, "tr"):
                columns = row.find_elements(By.TAG_NAME, "td")
                if len(columns) == 2:
                    key = columns[0].text.strip()
                    value = columns[1].text.strip()
                    if key and value:
                        table_data.append({"Key": key, "Value": value})
            if table_data:
                specifications["Specs"].append({"Header": title, "Attributes": table_data})
        if not specifications["Specs"]:
            specifications["Specs"] = "No specifications found"
    except Exception as e:
        logger.error(f"❌ Error retrieving specifications: {e}")
        specifications = {"Specs": "Error retrieving specifications"}
    return json.dumps(specifications, indent=4)

# == Download and save images ==
def download_image(image_url, product_mpn, img_count=None, image_type="price"):
    try:
        img_data = requests.get(image_url).content
        img_folder = "product_images"
        os.makedirs(img_folder, exist_ok=True)
        if img_count is None:
            img_name = f"{product_mpn}-{image_type}.jpg".lower()  # ✅ Correct: NX.KTDEK.002-price.jpg
        else:
            img_name = f"{product_mpn}-{img_count}-{image_type}.jpg".lower()  # ✅ Correct: NX.KTDEK.002-1-price.jpg
        img_path = os.path.join(img_folder, img_name)
        with open(img_path, 'wb') as f:
            f.write(img_data)
        logger.info(f"✅ Downloaded: {img_name}")
        return img_name
    except Exception as e:
        logger.error(f"❌ Failed to download image: {e}")
        return None

# == Tag's Function ==
def scrape_tags(driver):
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
        logger.error(f"Error retrieving tags: {e}")
        return {"Tags": "Error retrieving tags"}

# == Key Feature's ==
def scrape_key_features(driver):
    key_features = []
    try:
        features_div_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section/div/div[1]/div[2]/div[3]/div[2]'
        wait_for_element(driver, features_div_xpath)
        feature_items_xpath = features_div_xpath + '/ul/li'
        feature_elements = driver.find_elements(By.XPATH, feature_items_xpath)
        for el in feature_elements:
            text = el.text.strip()
            if text:
                key_features.append(text)
        if not key_features:
            key_features = ["N/A"]
        return json.dumps({"Key_Feature": key_features}, indent=4)
    except Exception as e:
        print(f"❌ Exception while scraping key features: {e}")
        return json.dumps({"Key_Feature": ["Error retrieving features"]}, indent=4)

# == FAQ's ==
def scrape_faqs(driver):
    faqs = []
    try:
        faq_section_xpath = '//*[@id="maincontent"]/app-dynamic-page/app-pdp/section[3]'
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, faq_section_xpath))
        )
        faq_heading = driver.find_element(By.XPATH, faq_section_xpath)
        driver.execute_script("arguments[0].scrollIntoView(true);", faq_heading)
        time.sleep(2)
        tabs_xpath = "//p-accordiontab//a"
        tabs = driver.find_elements(By.XPATH, tabs_xpath)
        for tab in tabs:
            driver.execute_script("arguments[0].click();", tab)
            time.sleep(2)
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        all_sections = soup.find_all('p-accordiontab')
        ignore_titles = {"Product Overview", "Specifications", "Customer Reviews","From Manufacturer"}
        for tab in all_sections:
            title = tab.find('span', class_='p-accordion-header-text')
            answer_block = tab.find('div', {'role': 'region'})
            if title and answer_block:
                q = title.get_text(strip=True)
                a = answer_block.get_text(strip=True)
                if q not in ignore_titles:
                    faqs.append({"Question": q, "Answer": a})
        if not faqs:
            faqs.append({"Question": "N/A", "Answer": "No FAQs found"})
        return {"FAQs": faqs}
    except Exception as e:
        logger.error(f"Error locating FAQ section: {e}")
        return {"FAQs": [{"Question": "N/A", "Answer": "FAQ section not found"}]}

#  == Check Product Valid or Not ==
def validate_product_link(driver, product_link):
    """
    Validates a product URL by checking for the presence of the 'Product Overview' section.
    Returns True if valid, else False.
    """
    try:
        driver.get(product_link)
        handle_cookie_popup(driver)
        handle_poster_popup(driver)
        logger.info(f"🔎 Validating: {product_link}")
        try:
            wait_for_element(driver, '//*[@id="maincontent"]/app-dynamic-page')
            logger.info("✅ Main content loaded")
        except Exception as e:
            logger.error(f"❌ Main content not loaded: {e}")
            return False 
        product_overview_xpath = '//*[contains(text(), "Product Overview")]'
        try:
            WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, product_overview_xpath))
            )
            logger.info("✅ Product Overview section found.")
            return True
        except:
            logger.warning("❌ Product Overview section not found.")
            return False
    except Exception as e:
        logger.error(f"❌ Exception during validation: {e}")
        return False
    
def scrape_with_retries(product_link):
    for attempt in range(1, 2):  # Max 2 retries
        logger.info(f"🔁 Attempt {attempt} for {product_link}")
        result = scrape_product_page(product_link)
        if result:
            return result
        time.sleep(1)  # Short pause before retrying
    logger.error(f"❌ Failed after 2 attempts: {product_link}")
    return None

# == Resume scraped links ==
scraped_links_file = "scraped_links.txt"
if os.path.exists(scraped_links_file):
    with open(scraped_links_file, "r") as f:
        already_scraped_links = set(line.strip() for line in f.readlines())
else:
    already_scraped_links = set()

# === File paths ===
invalid_links_file = "invalid_links.txt"
scraped_links_file = "scraped_links.txt"

# === Load previously scraped and invalid links ===
if os.path.exists(scraped_links_file):
    with open(scraped_links_file, "r") as f:
        already_scraped_links = set(line.strip() for line in f)
else:
    already_scraped_links = set()

if os.path.exists(invalid_links_file):
    with open(invalid_links_file, "r") as f:
        already_invalid_links = set(line.strip() for line in f)
else:
    already_invalid_links = set()

scraped_data = []
failed_links = []

# === Main scraping loop ===
for link in tqdm(df['Links'], desc="🔍 Scraping product pages"):
    print(f"\n🔄 Processing: {link}")
    if link in already_scraped_links:
        print("⏭️ Already scraped, skipping.")
        continue
    if link in already_invalid_links:
        print("⛔ Already marked invalid, skipping.")
        continue
    # Step 1: Validate product
    try:
        val_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=get_chrome_options())
        is_valid = validate_product_link(val_driver, link)
        val_driver.quit()
    except Exception as e:
        print(f"❌ Validation error for {link}: {e}")
        continue
    if not is_valid:
        print("❌ Invalid product. Marking as skipped.")
        with open(invalid_links_file, "a") as f:
            f.write(link + "\n")
        already_invalid_links.add(link)
        continue
    print("✅ Valid product confirmed. Scraping...")
    # Step 2: Scrape valid product
    try:
        product_data = scrape_with_retries(link)
        if product_data:
            scraped_data.append(product_data)
            logger.info(f"✅ Scraped successfully: {link}")
            with open(scraped_links_file, "a") as f:
                f.write(link + "\n")
        else:
            logger.warning(f"⚠️ Failed to scrape after retries: {link}")
            failed_links.append(link)
    except Exception as e:
        logger.error(f"❌ Scraping failed: {e}")
        failed_links.append(link)
# ✅ Final save
print(f"\n📦 Finished. Total products scraped: {len(scraped_data)}")
if scraped_data:
    current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_filename = f"scraped_product_data_{current_time}.xlsx"
    scraped_df = pd.DataFrame(scraped_data)
    scraped_df.to_excel(output_filename, index=False)
    logger.info(f"✅ Data saved to {output_filename}")
    print(f"✅ Data saved to Excel: {output_filename}")
else:
    print("⚠️ No products were scraped. Excel file not created.")

with open(scraped_links_file, "a") as f:
    for item in scraped_data:
        f.write(item['Link'] + "\n")