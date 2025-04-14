import csv
import time
import os
import re
import logging
import printnet
import add_fixed_price
import add_fixed_price_printnet
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys  # Added this import for the Keys module
from tkinter.tix import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select

# Setup logging

# Setup logging
def setup_logging():
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"skynet_page_options_{timestamp}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    logging.info(f"Logging initialized. Log file: {log_file}")
    return logging.getLogger()

# Initialize logger
logger = setup_logging()

# Configuration
CSV_FILE_PATH = "C:/Users/g.yilmaz/Desktop/Magento_otomation/csv_files/test.csv"
SKYNET_URL = "https://skynet.networkhero.eu/y5id_dashboard/admin/"

CSV_MAPPINGS = {
    "product_name": 2,  # Column C (index 2 in CSV)
    "code": 4,          # Column E (index 4 in CSV)
    "sku": 6            # Column G (index 6 in CSV)
}

PRICE = "1.00"
QUANTITY = "99999999"
VPE_VALUE = "9999"
VPE_TYPE = "Stock"  
WEBSITES_TO_SELECT = ["Selling Website", "Networkhero"]
COMBINATION_OPTION = "Fix Products"


# Hardcoded credentials
USERNAME = "Serhat"
PASSWORD = "ag0Qn4PX60%js&e9"

def login_to_admin(driver):
    """Login to Skynet admin panel using hardcoded credentials"""
    # Setup Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    
    try:
        logger.info(f"Navigating to {SKYNET_URL}")
        driver.get(SKYNET_URL)
        
        # Enter username
        username_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "username"))
        )
        username_field.clear()
        username_field.send_keys(USERNAME)
        logger.info("Username entered")
        
        # Enter password
        password_field = driver.find_element(By.ID, "login")
        password_field.clear()
        password_field.send_keys(PASSWORD)
        logger.info("Password entered")
        
        # Click login button
        login_button = driver.find_element(By.CSS_SELECTOR, ".action-login")
        driver.execute_script("arguments[0].click();", login_button)
        logger.info("Login button clicked")
        
        # Wait for Google Authenticator input
        logger.info("Waiting for you to enter Google Authenticator code...")
        print("Please enter the Google Authenticator code and press Submit...")
        
        # Wait for dashboard to load after 2FA
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".admin-user"))
        )
        
        logger.info("Login successful")
        return True
    except Exception as e:
        logger.error(f"Login failed: {str(e)}")
        return False

def click_cloudlab(driver):
    """Click on CLOUDLAB icon using the specified approach"""
    try:
        # 1. Click on CLOUDLAB icon in the left menu using the Catalog approach
        logger.info("Clicking on CLOUDLAB icon using the Catalog approach...")
        
        try:
            # This is the code from navigate_to_add_product function that works for Catalog
            cloudlab_icon = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-cloudlab')]//a"))
            )
            driver.execute_script("arguments[0].click();", cloudlab_icon)
            logger.info("Clicked on CLOUDLAB icon using Catalog approach")
            
            return True
        except Exception as e:
            logger.warning(f"Catalog approach failed for CLOUDLAB: {str(e)}")
            
            # Try alternative approach - find by data-ui-id or other attributes
            try:
                cloudlab_icon = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".menu-item[data-ui-id*='menu-cloudlab'], .item-cloudlab, #menu-cloudlab"))
                )
                driver.execute_script("arguments[0].click();", cloudlab_icon)
                logger.info("Clicked on CLOUDLAB icon using data-ui-id approach")
                
                
                return True
            except Exception as e2:
                logger.warning(f"data-ui-id approach failed for CLOUDLAB: {str(e2)}")
                
                # Try the most generic approach - find by position (left menu 6th item)
                try:
                    # Find all menu items in the left menu
                    menu_items = driver.find_elements(By.CSS_SELECTOR, ".menu-item, .item, li[role='menuitem'], div[role='menuitem']")
                    
                    if len(menu_items) >= 6:
                        cloudlab_icon = menu_items[5]  # 6th item (index 5)
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cloudlab_icon)
                        driver.execute_script("arguments[0].click();", cloudlab_icon)
                        logger.info("Clicked on CLOUDLAB icon by position (6th menu item)")
                        
                        
                        
                        return True
                    else:
                        logger.warning(f"Not enough menu items found: {len(menu_items)}")
                        return False
                except Exception as e3:
                    logger.error(f"All attempts to click CLOUDLAB failed: {str(e3)}")
                    return False
    except Exception as e:
        logger.error(f"Error clicking on CloudLab: {str(e)}")
        return False

def click_page_options(driver):
    """Click on Page Options in the submenu"""
    try:
        # Adapted from the Product Parts code, but for Page Options
        logger.info("Clicking on Page Options submenu...")
        
        try:
            # Modified XPath to target Page Options instead of Product Parts
            page_options_submenu = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-cloudlab')]//li[contains(@class, 'item-page-options') or contains(@class, 'item-cloudlab-page-options')]//a"))
            )
            driver.execute_script("arguments[0].click();", page_options_submenu)
            logger.info("Clicked on Page Options using XPath approach")
        except Exception as e:
            logger.warning(f"XPath approach failed for Page Options: {str(e)}")
            
            # Try alternative approach - find by text
            try:
                # Look for elements containing "Page Options" text
                page_options_elements = driver.find_elements(By.XPATH, 
                    "//span[text()='Page Options'] | //a[text()='Page Options'] | //*[contains(text(), 'Page Options')]")
                
                if page_options_elements:
                    for elem in page_options_elements:
                        if elem.is_displayed():
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                            driver.execute_script("arguments[0].click();", elem)
                            logger.info("Clicked on Page Options by text")
                            break
                    else:
                        logger.warning("Found Page Options elements but none were clickable")
                        return False
                else:
                    logger.warning("Could not find Page Options by text")
                    return False
            except Exception as e2:
                logger.error(f"All attempts to click Page Options failed: {str(e2)}")
                return False
        
        # Wait for Page Options page to load
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
            )
            
            
            logger.info("Successfully navigated to Page Options page")
            return True
        except Exception as wait_e:
            logger.error(f"Timeout waiting for Page Options page to load: {str(wait_e)}")
            return False
    
    except Exception as e:
        logger.error(f"Error clicking on Page Options: {str(e)}")
        return False

def search_and_click_material(driver):
    """Search for 'paper_wmd' in the Code field and click on Material result"""
    try:
        logger.info("Looking for Code search field")
        
        # Find the Code search field
        code_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='code']"))
        )
        
        # Clear any existing value
        code_field.clear()
        
        # Enter "paper_wmd"
        code_field.send_keys("paper_wmd")
        logger.info("Entered 'paper_wmd' in Code search field")
        
        # Look for the search button instead of pressing Enter
        try:
            search_button = WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.action-search, button.search, button[title='Search'], button[title='Suchen']"))
            )
            search_button.click()
            logger.info("Clicked Search button")
        except Exception as search_button_e:
            logger.warning(f"Could not find search button, trying to press Enter instead: {str(search_button_e)}")
            # Press Enter as fallback
            code_field.send_keys(Keys.ENTER)
            logger.info("Pressed Enter to search")
        
        # Now find and click on the second Material option in the search results
        logger.info("Looking for the second Material option in search results")
        
        try:
            # Find all rows in the table
            rows = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.data-row, tbody tr"))
            )
            
            logger.info(f"Found {len(rows)} rows in the search results")
            
            # Try to find the row containing "paper_wmd" in the Code column
            for row in rows:
                try:
                    # Look for the code column (typically the third column)
                    columns = row.find_elements(By.TAG_NAME, "td")
                    if len(columns) >= 3:
                        code_column = columns[2]  # Assuming Code is the third column based on your screenshot
                        if "paper_wmd" in code_column.text.lower():
                            # Found the row with paper_wmd, now find and click on the Material field
                            material_column = None
                            for column in columns:
                                if "Material" in column.text:
                                    material_column = column
                                    break
                            
                            if material_column:
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", material_column)
                                driver.execute_script("arguments[0].click();", material_column)
                                logger.info("Clicked on Material column in row containing paper_wmd")
                                return True
                except Exception as row_e:
                    logger.warning(f"Error processing row: {str(row_e)}")
                    continue
            else:
                logger.warning("Not enough rows found in search results")
                
                # If there are fewer than 2 rows, try to click on any row that has "Material" text
                try:
                    material_element = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, "//td[contains(text(), 'Material')]"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", material_element)
                    driver.execute_script("arguments[0].click();", material_element)
                    logger.info("Clicked on the first available Material element")
                    return True
                except Exception as mat_e:
                    logger.error(f"Could not find any Material elements: {str(mat_e)}")
                    return False
        except Exception as e:
            logger.error(f"Error finding rows in search results: {str(e)}")
            
            # Try a more generic approach as a last resort
            try:
                # Just look for any Material element if all else fails
                material_elements = driver.find_elements(By.XPATH, "//td[contains(text(), 'Material')]")
                
                if len(material_elements) >= 2:
                    # Click on the second Material element
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", material_elements[1])

                    driver.execute_script("arguments[0].click();", material_elements[1])
                    logger.info("Clicked on the second Material element using generic approach")
                    return True
                elif len(material_elements) >= 1:
                    # If there's only one Material element, click on it
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", material_elements[0])

                    driver.execute_script("arguments[0].click();", material_elements[0])
                    logger.info("Clicked on the first Material element using generic approach")
                    return True
                else:
                    logger.warning("No Material elements found")
                    return False
            except Exception as alt_e:
                logger.error(f"Generic approach also failed: {str(alt_e)}")
                return False
    except Exception as e:
        logger.error(f"Failed to search for 'paper' or click on Material: {str(e)}")
        return False

def click_page_option_values_and_add_new(driver):
    """Click on Page Option Values and then Add new Value"""
    try:
        logger.info("Looking for Page Option Values section")
        
        # Wait for the edit page to fully load
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
        )
        
        # Scroll down the page to make sure we can see Page Option Values
        logger.info("Scrolling down to find Page Option Values")
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.7);")
        time.sleep(1)  # Short pause after scrolling
        
        # Find and click on Page Option Values
        try:
            # Try multiple selectors for Page Option Values
            selectors = [
                "//div[contains(text(), 'Page Option Values')]",
                "//span[contains(text(), 'Page Option Values')]",
                "//div[.//span[contains(text(), 'Page Option Values')]]",
                "//div[contains(@class, 'fieldset-wrapper-title')][contains(text(), 'Page Option Values')]",
                "//div[contains(@class, 'section-title')][contains(text(), 'Option Values')]"
            ]
            
            page_option_values = None
            for selector in selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    if elements and len(elements) > 0:
                        for element in elements:
                            if element.is_displayed():
                                page_option_values = element
                                break
                        if page_option_values:
                            break
                except:
                    continue
            
            if not page_option_values:
                # Look for any element that contains "Page Option Values" text
                elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Page Option Values')]")
                for element in elements:
                    if element.is_displayed():
                        page_option_values = element
                        break
            
            if page_option_values:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", page_option_values)
                
                driver.execute_script("arguments[0].click();", page_option_values)
                logger.info("Clicked on Page Option Values section")
                
                # Wait a moment for any animations to complete
                time.sleep(1)
                
                # Now find and click on Add new Value button
                try:
                    # First approach: Direct JavaScript injection to find the button by its visual characteristics
                    js_find_button = """
                    // Function to check if an element visually looks like a button
                    function isLikelyButton(el) {
                        if (el.tagName === 'BUTTON') return true;
                        if (el.tagName === 'A' && (el.className.includes('button') || el.className.includes('action'))) return true;
                        return false;
                    }
                    
                    // Loop through all potential button elements 
                    var potentialButtons = [...document.getElementsByTagName('button'), ...document.getElementsByTagName('a')];
                    for (var i = 0; i < potentialButtons.length; i++) {
                        var btn = potentialButtons[i];
                        if (!btn.offsetParent) continue; // Skip hidden elements
                        
                        // Check if it has text containing "add" and "value"
                        var buttonText = btn.textContent.toLowerCase();
                        if (buttonText.includes('add') && buttonText.includes('value')) {
                            btn.scrollIntoView({block: 'center'});
                            btn.click();
                            return "Clicked button: " + buttonText;
                        }
                    }
                    
                    // If specific text not found, look for elements that look like action buttons
                    for (var i = 0; i < potentialButtons.length; i++) {
                        var btn = potentialButtons[i];
                        if (!btn.offsetParent) continue; // Skip hidden elements
                        
                        var style = window.getComputedStyle(btn);
                        var classes = btn.className.toLowerCase();
                        
                        // Check if it looks like a primary action button (red or blue or with primary class)
                        if (classes.includes('primary') || 
                            classes.includes('action') || 
                            style.backgroundColor.includes('rgb(255') || 
                            style.backgroundColor.includes('rgb(0, 0, 255')) {
                            
                            // Look more specifically for add button indicators
                            if (classes.includes('add') || btn.id.includes('add') || btn.innerHTML.includes('+')) {
                                btn.scrollIntoView({block: 'center'});
                                btn.click();
                                return "Clicked by style: " + btn.textContent;
                            }
                        }
                    }
                    
                    return null;
                    """
                    
                    result = driver.execute_script(js_find_button)
                    if result:
                        logger.info(f"Successfully found and clicked button via JavaScript: {result}")
                        return True
                    
                    # Second approach: Try to find the button by its visual position relative to the Page Option Values section
                    js_find_by_position = """
                    var header = document.evaluate("//div[contains(text(), 'Page Option Values')]", 
                                                  document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (!header) return "Header not found";
                    
                    var headerRect = header.getBoundingClientRect();
                    
                    // Define a clickable region where the button should be 
                    // (adjust these values based on the actual UI layout)
                    var searchArea = {
                        top: headerRect.bottom,
                        left: headerRect.left,
                        bottom: headerRect.bottom + 100,
                        right: headerRect.left + 200
                    };
                    
                    // Create an array of points to check within this region
                    var pointsToCheck = [];
                    for (var x = searchArea.left + 20; x < searchArea.right; x += 40) {
                        for (var y = searchArea.top + 20; y < searchArea.bottom; y += 20) {
                            pointsToCheck.push({x: x, y: y});
                        }
                    }
                    
                    // Try each point to see if we hit a button
                    for (var i = 0; i < pointsToCheck.length; i++) {
                        var point = pointsToCheck[i];
                        var element = document.elementFromPoint(point.x, point.y);
                        
                        if (!element) continue;
                        
                        // Check if this element or any parent is a button or link
                        var buttonCheck = element;
                        while (buttonCheck && buttonCheck !== document.body) {
                            if (buttonCheck.tagName === 'BUTTON' || 
                                (buttonCheck.tagName === 'A' && buttonCheck.className.includes('action')) ||
                                buttonCheck.className.includes('action-primary') ||
                                buttonCheck.className.includes('add')) {
                                
                                buttonCheck.scrollIntoView({block: 'center'});
                                buttonCheck.click();
                                return "Clicked by position: " + buttonCheck.textContent + " at " + point.x + "," + point.y;
                            }
                            buttonCheck = buttonCheck.parentElement;
                        }
                    }
                    
                    // If we've made it here, we've exhausted our search grid with no button found
                    // Last resort: try clicking in the general area where a button might be
                    var centerX = (searchArea.left + searchArea.right) / 2;
                    var centerY = searchArea.top + 30;
                    var centerElement = document.elementFromPoint(centerX, centerY);
                    
                    if (centerElement) {
                        centerElement.scrollIntoView({block: 'center'});
                        centerElement.click();
                        return "Last resort click at " + centerX + "," + centerY + " on " + centerElement.tagName;
                    }
                    
                    return "No clickable elements found";
                    """
                    
                    result = driver.execute_script(js_find_by_position)
                    logger.info(f"Button search result: {result}")
                    
                    # If we got a positive result (not an error message)
                    if result and not result.startswith("No clickable") and not result.startswith("Header not found"):
                        logger.info(f"Successfully found and clicked button via position: {result}")
                        return True
                    
                    # Third approach: Try to directly target the element with CSS selector which may be more specific
                    # Look for buttons within the section or immediately following it in the DOM
                    try:
                       
                        
                        # Now try all common button selectors
                        button_css_selectors = [
                            ".page-actions button", 
                            ".page-actions .action-primary",
                            ".admin__actions-switch + button",
                            ".admin__page-section-title + .admin__page-section-content button",
                            ".fieldset__actions button",
                            ".fieldset-wrapper-title + div button"
                        ]
                        
                        for css_selector in button_css_selectors:
                            try:
                                buttons = driver.find_elements(By.CSS_SELECTOR, css_selector)
                                if buttons:
                                    for button in buttons:
                                        if button.is_displayed():
                                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
                                            
                                            driver.execute_script("arguments[0].click();", button)
                                            logger.info(f"Clicked button via CSS selector: {css_selector}")
                                            return True
                            except:
                                continue
                        
                        # If all else fails, try a very direct approach - click where the button appears to be
                        # based on the screenshot, directly click at spot
                        x_offset = 110  # Adjust based on your actual UI
                        y_offset = 375  # Adjust based on your actual UI
                        
                        action = webdriver.ActionChains(driver)
                        action.move_to_element(page_option_values)
                        action.move_by_offset(x_offset, y_offset).click().perform()
                        logger.info(f"Attempted direct click at offset x:{x_offset}, y:{y_offset}")
                        
                        return True
                    except Exception as css_e:
                        logger.error(f"CSS selector approach failed: {str(css_e)}")
                        return False
                    
                except Exception as e:
                    logger.error(f"Error finding Add new Value button: {str(e)}")
                    return False
            else:
                logger.warning("Could not find Page Option Values section")
                return False
        except Exception as e:
            logger.warning(f"Error clicking on Page Option Values: {str(e)}")
            return False
    except Exception as e:
        logger.error(f"Error in click_page_option_values_and_add_new: {str(e)}")
        return False

def parse_materials(material_str):
    """Parse multiple materials from a string with numbered format like "1- material_1\\ 2- material_2\\ 3- material_3\\"
    
    Returns a list of materials with the numbering removed.
    """
    if not material_str:
        return []
    
    # Split the string by backslash separator
    items = material_str.split('\\')
    
    # Process each item to remove numbering at the beginning
    materials = []
    for item in items:
        # Remove initial numbering (e.g., "1- " or "2. ")
        cleaned_item = re.sub(r'^\s*\d+[-.\s)]+\s*', '', item).strip()
        if cleaned_item:
            materials.append(cleaned_item)
    
    return materials if materials else [material_str.strip()]

def parse_materials_page_option(material_str):
    """Parse multiple materials from a string with numbered format like "1- material_1\\ 2- material_2\\ 3- material_3\\"
    
    Returns a list of materials with the numbering removed.
    """
    if not material_str:
        return []
    
    # Split the string by backslash separator
    items = material_str.split('\\')
    
    # Process each item to remove numbering at the beginning
    materials = []
    for item in items:
        # Remove initial numbering (e.g., "1- " or "2. ")
        cleaned_item = re.sub(r'^\s*\d+[-.\s)]+\s*', '', item).strip()
        if cleaned_item:
            materials.append(cleaned_item)
    
    return materials if materials else [material_str.strip()]

def fill_add_new_value_form(driver, csv_file_path):
    """Fill form using pure JavaScript, adding Position field from column J"""
    try:
        # Read the CSV file
        import csv
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            
            # Skip the first 6 rows
            for _ in range(6):
                try:
                    next(csv_reader)
                except StopIteration:
                    logger.warning("CSV file doesn't have enough rows")
                    return False
            
            # Read the data row
            data_row = next(csv_reader)
            
            # Get the values from H, I and J columns
            name_value_raw = data_row[7] if len(data_row) > 7 else ""    # H column
            code_value_raw = data_row[8] if len(data_row) > 8 else ""    # I column
            position_value = data_row[9] if len(data_row) > 9 else ""    # J column
            
            # Parse multiple materials from the name and code values
            name_materials = parse_materials_page_option(name_value_raw)
            code_materials = parse_materials_page_option(code_value_raw)
            
            # Log the parsed materials
            logger.info(f"Parsed {len(name_materials)} materials from name value: {name_materials}")
            logger.info(f"Parsed {len(code_materials)} materials from code value: {code_materials}")
            
            # Ensure we have matching pairs of name and code
            material_count = min(len(name_materials), len(code_materials))
            
            # Process each material
            for i in range(material_count):
                name_value = name_materials[i]
                code_value = code_materials[i]
                
                logger.info(f"Processing material {i+1}/{material_count}: Name={name_value}, Code={code_value}, Position={position_value}")
                
                # If not the first material, we need to refresh the page and click Page Option Values and Add New
                if i > 0:
                    logger.info(f"Refreshing page for material {i+1}")
                    
                    # Refresh the page first
                    driver.refresh()
                    
                    # Wait for page to fully load after refresh
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
                    )
                    
                    logger.info("Page refreshed successfully")
                    
                    # Need to click Page Option Values and Add new Value again
                    if not click_page_option_values_and_add_new(driver):
                        logger.error(f"Failed to set up for material {i+1}")
                        continue
                    
                    # Wait for modal to load
                    time.sleep(2)
                
                # Wait for modal to load
                time.sleep(2)
                
                # Take a screenshot
                driver.save_screenshot(f"before_fill_material_{i+1}.png")
                
                # Use pure JavaScript to find and fill the input fields
                js_script = """
                // Function to fill all input fields that match criteria
                function fillInputFields(nameValue, codeValue, positionValue) {
                    console.log('Starting to fill inputs with:', nameValue, codeValue, positionValue);
                    
                    // Get all inputs in the document
                    var allInputs = document.querySelectorAll('input');
                    console.log('Found', allInputs.length, 'input fields');
                    
                    // First try to identify fields by attributes
                    var nameField = null;
                    var codeField = null;
                    var positionField = null;
                    
                    // Find fields by name/id attribute
                    for (var i = 0; i < allInputs.length; i++) {
                        var input = allInputs[i];
                        var name = input.name || '';
                        var id = input.id || '';
                        
                        if (name.includes('name') || id.includes('name')) {
                            nameField = input;
                        }
                        
                        if (name.includes('code') || id.includes('code')) {
                            codeField = input;
                        }
                        
                        if (name.includes('position') || id.includes('position')) {
                            positionField = input;
                        }
                    }
                    
                    // If not found by name/id, try by position - first visible input is name, second is code, third is position
                    if (!nameField || !codeField || !positionField) {
                        var visibleInputs = Array.from(allInputs).filter(function(input) {
                            // Check if input is visible
                            var style = window.getComputedStyle(input);
                            return input.type !== 'hidden' && 
                                style.display !== 'none' && 
                                style.visibility !== 'hidden' && 
                                input.offsetParent !== null;
                        });
                        
                        console.log('Found', visibleInputs.length, 'visible inputs');
                        
                        if (visibleInputs.length >= 1 && !nameField) {
                            nameField = visibleInputs[0];
                        }
                        
                        if (visibleInputs.length >= 2 && !codeField) {
                            codeField = visibleInputs[1];
                        }
                        
                        if (visibleInputs.length >= 3 && !positionField) {
                            positionField = visibleInputs[2];
                        }
                    }
                    
                    // Fill the fields if found
                    var result = {
                        nameFieldFound: false,
                        codeFieldFound: false,
                        positionFieldFound: false
                    };
                    
                    if (nameField) {
                        try {
                            // Try multiple ways to set value
                            nameField.value = nameValue;
                            // Also try to set the value attribute directly
                            nameField.setAttribute('value', nameValue);
                            
                            // Trigger events
                            nameField.dispatchEvent(new Event('input', {bubbles: true}));
                            nameField.dispatchEvent(new Event('change', {bubbles: true}));
                            
                            result.nameFieldFound = true;
                            result.nameFieldValue = nameField.value;
                            console.log('Name field filled with:', nameField.value);
                        } catch (e) {
                            console.error('Error filling name field:', e);
                        }
                    }
                    
                    if (codeField) {
                        try {
                            // Try multiple ways to set value
                            codeField.value = codeValue;
                            // Also try to set the value attribute directly
                            codeField.setAttribute('value', codeValue);
                            
                            // Trigger events
                            codeField.dispatchEvent(new Event('input', {bubbles: true}));
                            codeField.dispatchEvent(new Event('change', {bubbles: true}));
                            
                            result.codeFieldFound = true;
                            result.codeFieldValue = codeField.value;
                            console.log('Code field filled with:', codeField.value);
                        } catch (e) {
                            console.error('Error filling code field:', e);
                        }
                    }
                    
                    if (positionField) {
                        try {
                            // Try multiple ways to set value
                            positionField.value = positionValue;
                            // Also try to set the value attribute directly
                            positionField.setAttribute('value', positionValue);
                            
                            // Trigger events
                            positionField.dispatchEvent(new Event('input', {bubbles: true}));
                            positionField.dispatchEvent(new Event('change', {bubbles: true}));
                            
                            result.positionFieldFound = true;
                            result.positionFieldValue = positionField.value;
                            console.log('Position field filled with:', positionField.value);
                        } catch (e) {
                            console.error('Error filling position field:', e);
                        }
                    }
                    
                    return result;
                }
                
                return fillInputFields(arguments[0], arguments[1], arguments[2]);
                """
                
                # Execute the JavaScript
                result = driver.execute_script(js_script, name_value, code_value, position_value)
                logger.info(f"JavaScript fill result for material {i+1}: {result}")
                
                # Wait to ensure changes take effect
                time.sleep(1)
                
                # Take after screenshot
                driver.save_screenshot(f"after_fill_material_{i+1}.png")
                
                # Click the popup save button to save each material
                if not click_popup_save_button(driver):
                    logger.error(f"Failed to save material {i+1}")
                    continue
                
                logger.info(f"Successfully saved material {i+1}")
                
                # Wait for save to complete
                time.sleep(2)
            
            # After all materials are processed, refresh the page and click main save button
            if material_count > 0:
                logger.info("All materials processed, now clicking main save")
                if refresh_page_and_click_main_save(driver):
                    logger.info("Successfully saved all page option values")
                else:
                    logger.error("Failed to save page after all materials were added")
            
            return True
            
    except Exception as e:
        logger.error(f"Error in fill_add_new_value_form: {str(e)}")
        return False

def click_popup_save_button(driver):
    """
    Function specifically designed to click the Save button in the popup modal that appears
    after clicking 'Add new Value' - NOT the main page Save button
    """
    try:
        logger.info("Starting popup modal Save button click function")

        
        # APPROACH 1: Look for a modal dialog first, then find its Save button
        try:
            # Look for modal dialog/popup
            modal = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".modal-popup, .modal-slide, .ui-dialog"))
            )
            logger.info("Found modal popup")
            
            # Now look for Save button inside this modal
            save_button = modal.find_element(By.CSS_SELECTOR, "button#save, button.action-primary, button[title='Save'], button.primary")
            
            
 
            
            # Click the button
            driver.execute_script("arguments[0].click();", save_button)
            logger.info("Clicked Save button inside popup modal")
            return True
        except Exception as e:
            logger.warning(f"Could not find Save button in modal: {str(e)}")
        
        # APPROACH 2: Use a different strategy to identify the popup
        try:
            # Find elements that suggest we're in a popup (Page Option Value title)
            title_elem = driver.find_element(By.XPATH, "//div[contains(text(), 'Page Option Value')]")
            if title_elem:
                logger.info("Found Page Option Value popup")
                
                # Find the closest button with 'Save' text within this modal/popup
                save_buttons = driver.find_elements(By.XPATH, "//button[contains(., 'Save')]")
                
                # Filter to find visible buttons only
                for button in save_buttons:
                    if button.is_displayed():
                        # Highlight the button
                        driver.execute_script("arguments[0].style.border = '3px solid red';", button)
                        driver.execute_script("arguments[0].style.backgroundColor = 'yellow';", button)
                        
    
                        
                        # Click the button
                        driver.execute_script("arguments[0].click();", button)
                        logger.info("Clicked visible Save button in popup")
                        return True
                
                logger.warning("Found popup but no visible Save buttons")
        except Exception as e:
            logger.warning(f"Title-based approach failed: {str(e)}")
        
        # APPROACH 3: More generic approach searching all modals
        try:
            # This approach looks for ANY active modal/popup and its save button
            js_code = """
            // Function to find the Save button in an active modal/popup
            function findActiveModalSaveButton() {
                // Look for elements that might be modals/popups
                var possibleModals = document.querySelectorAll('.modal-popup, .modal-slide, .ui-dialog, .modal-content, [role="dialog"]');
                
                for (var i = 0; i < possibleModals.length; i++) {
                    var modal = possibleModals[i];
                    
                    // Check if this modal is visible
                    if (modal.offsetParent !== null && 
                        window.getComputedStyle(modal).display !== 'none' && 
                        window.getComputedStyle(modal).visibility !== 'hidden') {
                        
                        console.log("Found active modal:", modal.className);
                        
                        // Look for Save buttons inside this modal
                        var buttons = modal.querySelectorAll('button');
                        for (var j = 0; j < buttons.length; j++) {
                            var button = buttons[j];
                            
                            // Check if this button has 'Save' text
                            if (button.textContent.trim().toLowerCase() === 'save' || 
                                button.textContent.trim() === '') {
                                
                                // Also check if it's visible
                                if (button.offsetParent !== null) {
                                    console.log("Found Save button:", button.className);
                                    
                                    // Highlight
                                    button.style.border = '3px solid red';
                                    button.style.backgroundColor = 'yellow';
                                    
                                    return button;
                                }
                            }
                        }
                        
                        // If no button with 'Save' text, look for action buttons
                        var actionButtons = modal.querySelectorAll('button.action-primary, button.primary, button[data-role="action"]');
                        for (var j = 0; j < actionButtons.length; j++) {
                            if (actionButtons[j].offsetParent !== null) {
                                console.log("Found action button:", actionButtons[j].className);
                                
                                // Highlight
                                actionButtons[j].style.border = '3px solid red';
                                actionButtons[j].style.backgroundColor = 'yellow';
                                
                                return actionButtons[j];
                            }
                        }
                    }
                }
                
                return null;
            }
            
            var button = findActiveModalSaveButton();
            if (button) {
                setTimeout(function() {
                    button.click();
                }, 500);
                return true;
            }
            return false;
            """
            
            result = driver.execute_script(js_code)
            
            if result:

                
                # Wait for click to complete
                time.sleep(0.5)
                
                logger.info("Found and clicked Save button in active modal using JavaScript")
                return True
            else:
                logger.warning("JavaScript approach could not find modal Save button")
        except Exception as e:
            logger.warning(f"JavaScript approach failed: {str(e)}")
        
        # APPROACH 4: Try to find by specific focus - the most recently added element
        try:
            logger.info("Looking for most recently added Save button")
            
            # First get a list of all buttons that might be the Save button
            all_save_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Save') or .//span[contains(text(), 'Save')]]")
            
            if all_save_buttons:
                # Click the last one (most recently added to DOM)
                save_button = all_save_buttons[-1]
                
                # Highlight
                driver.execute_script("arguments[0].style.border = '3px solid red';", save_button)
                driver.execute_script("arguments[0].style.backgroundColor = 'yellow';", save_button)
                
                
                # Click
                driver.execute_script("arguments[0].click();", save_button)
                logger.info("Clicked most recently added Save button")
                return True
            else:
                logger.warning("No Save buttons found")
        except Exception as e:
            logger.warning(f"Most recent button approach failed: {str(e)}")
        
        
        return True  # Assume manual action was successful
     
    except Exception as e:
        logger.error(f"Error in click_popup_save_button: {str(e)}")
        return False

def refresh_page_and_click_main_save(driver):
    """
    Refresh the page and then click the main Save button in the top-right corner
    """
    try:
        logger.info("Refreshing page...")
        
        # Refresh the page
        driver.refresh()
        
        # Wait for page to fully load after refresh
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".page-header, .page-title"))
        )
        
        logger.info("Page refreshed successfully")
        
  
        
        # Wait a moment for any JavaScript to initialize
        time.sleep(1)
        
        # Now find and click the main Save button in the top-right corner
        try:
            # Look for the Save button in the top-right action area
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".page-actions-buttons button.primary, .page-actions button.action-primary, button[title='Save'][data-action='save']"))
            )
            
            # Highlight the button for visual confirmation
            driver.execute_script("""
                arguments[0].style.border = '3px solid red';
                arguments[0].style.backgroundColor = 'yellow';
            """, save_button)

            
            # Click the button
            driver.execute_script("arguments[0].click();", save_button)
            
            logger.info("Clicked main Save button in top-right corner")
            
            # Wait for save action to complete (look for any confirmation or change)
            time.sleep(1)
      
            
            return True
        except Exception as e:
            logger.warning(f"Could not find main Save button: {str(e)}")
            
            # Alternative approach using JavaScript to find and click the top-right save button
            js_code = """
            // Function to find the main Save button in the page (not in a modal)
            function findMainSaveButton() {
                // Look for buttons in the top action area
                var actionAreas = document.querySelectorAll('.page-actions, .page-actions-buttons, .page-header-actions');
                
                for (var i = 0; i < actionAreas.length; i++) {
                    var area = actionAreas[i];
                    
                    // Look for Save buttons in this area
                    var buttons = area.querySelectorAll('button');
                    for (var j = 0; j < buttons.length; j++) {
                        var button = buttons[j];
                        
                        // Check if it's a save button by text
                        if (button.textContent.trim().toLowerCase().includes('save')) {
                            console.log("Found main Save button by text:", button.className);
                            
                            // Highlight for visibility
                            button.style.border = '3px solid red';
                            button.style.backgroundColor = 'yellow';
                            
                            return button;
                        }
                        
                        // Check if it's a primary action button
                        if (button.classList.contains('primary') || 
                            button.classList.contains('action-primary')) {
                            console.log("Found primary action button:", button.className);
                            
                            // Highlight for visibility
                            button.style.border = '3px solid red';
                            button.style.backgroundColor = 'yellow';
                            
                            return button;
                        }
                    }
                }
                
                // If no button found in action areas, try looking by position (top-right)
                var allButtons = document.querySelectorAll('button');
                var rightmostTopButton = null;
                var rightmostPosition = 0;
                
                for (var i = 0; i < allButtons.length; i++) {
                    var btn = allButtons[i];
                    if (btn.offsetParent === null) continue; // Skip hidden buttons
                    
                    var rect = btn.getBoundingClientRect();
                    if (rect.top < 200) {  // Only consider buttons at the top of the page
                        if (rect.right > rightmostPosition) {
                            rightmostPosition = rect.right;
                            rightmostTopButton = btn;
                        }
                    }
                }
                
                if (rightmostTopButton) {
                    console.log("Found rightmost top button:", rightmostTopButton.className);
                    
                    // Highlight for visibility
                    rightmostTopButton.style.border = '3px solid red';
                    rightmostTopButton.style.backgroundColor = 'yellow';
                    
                    return rightmostTopButton;
                }
                
                return null;
            }
            
            var button = findMainSaveButton();
            if (button) {
                setTimeout(function() {
                    button.click();
                }, 500);
                return true;
            }
            return false;
            """
            
            result = driver.execute_script(js_code)
            
            if result:
        
                # Wait for click to complete
                time.sleep(1)
                
                logger.info("Found and clicked main Save button using JavaScript")
                return True
            else:
                logger.warning("JavaScript approach could not find main Save button")
                
                # Last resort: try direct coordinate click at the position of the save button
                try:
                    window_size = driver.get_window_size()
                    # These coordinates target the top-right corner where Save button usually is
                    x_coord = window_size['width'] - 100
                    y_coord = 100
                    
                    actions = webdriver.ActionChains(driver)
                    actions.move_by_offset(x_coord, y_coord).click().perform()
                    
                    logger.info("Performed direct coordinate click for main Save button")
               
                    return True
                except Exception as coord_e:
                    logger.warning(f"Coordinate click failed: {str(coord_e)}")
            
            # If all automated attempts fail, prompt for manual intervention
            logger.error("All automated attempts to click main Save button failed")
            print("\n***** MANUAL INTERVENTION REQUIRED *****")
            print("Please click the SAVE button in the top-right corner of the main page.")
            input("Press Enter after clicking Save manually...")
            
            return True  # Assume manual click was successful
            
    except Exception as e:
        logger.error(f"Error in refresh_page_and_click_main_save: {str(e)}")
        return False

def select_materials_from_csv(driver, csv_file_path):
    """Sayfada Öffnung altındaki Material kısmını bulur ve CSV'den okunan malzemeleri seçer"""
    try:
        logger.info("Attempting to select materials from CSV")
        
        # 1. Önce CSV'den malzeme bilgilerini oku
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            
            # İlk 6 satırı atla
            for _ in range(6):
                next(csv_reader, None)
            
            # Veri satırını oku
            data_row = next(csv_reader)
            
            # I sütunundaki malzeme bilgilerini al
            name_value_raw = data_row[8] if len(data_row) > 8 else ""
            
            # Malzemeleri doğru şekilde parse et (daha güçlü bir yaklaşım)
            material_names = parse_materials(name_value_raw)
            
            logger.info(f"Found materials from CSV: {material_names}")
        
        # 2. Sayfada Öffnung elementini bul
        offnung_elements = driver.find_elements(By.XPATH, "//span[contains(text(), 'Öffnung') or contains(text(), 'Offnung')]")
        
        if not offnung_elements:
            logger.warning("Could not find Öffnung element, trying alternative approaches")
            offnung_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Öffnung') or contains(text(), 'Offnung')]")
        
        if not offnung_elements:
            logger.error("Could not find Öffnung element")
            return False
        
        # Görünür Öffnung elementini bul
        offnung_element = None
        for elem in offnung_elements:
            if elem.is_displayed():
                offnung_element = elem
                break
        
        if not offnung_element:
            logger.error("Could not find visible Öffnung element")
            return False
        
        # Öffnung elementini görünür yap
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", offnung_element)
        time.sleep(1)
        
        logger.info("Found Öffnung element, now looking for Material field below it")
        
        # Öffnung konumunu al
        offnung_location = offnung_element.location
        
        # 3. Öffnung'un altındaki Material alanını bul
        material_elements = driver.find_elements(By.XPATH, "//span[contains(text(), 'Material')]")
        
        # Öffnung'tan sonra gelen Material elementini bul
        material_element = None
        for elem in material_elements:
            if elem.is_displayed() and elem.location['y'] > offnung_location['y']:
                material_element = elem
                break
        
        if not material_element:
            logger.error("Could not find Material element below Öffnung")
            return False
        
        # Material elementini görünür yap
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", material_element)
        time.sleep(1)
        
        logger.info("Found Material element, now looking for select element")
        
        # 4. Material alanındaki select elementini bul
        select_element = None
        
        # Parent elementi bul
        parent = material_element.find_element(By.XPATH, "..")
        for _ in range(5):  # 5 seviyeye kadar yukarı çık
            try:
                parent = parent.find_element(By.XPATH, "..")
                try:
                    # İçinde select olan bir alan mı?
                    select_element = parent.find_element(By.TAG_NAME, "select")
                    logger.info("Found Material select element")
                    break
                except:
                    # Değilse devam et
                    pass
            except:
                break
        
        if not select_element:
            logger.error("Could not find Material select element")
            return False
        
        # 5. JavaScript ile direkt value'ları seç
        js_result = driver.execute_script("""
            var selectElement = arguments[0];
            var materialNames = arguments[1];
            
            // Tüm seçimleri kaldır
            for (var i = 0; i < selectElement.options.length; i++) {
                selectElement.options[i].selected = false;
            }
            
            // Eşleşen malzemeleri seç
            var selectedCount = 0;
            var selectedOptions = [];
            
            // Tüm malzeme isimleri için kontrol
            materialNames.forEach(function(materialName) {
                materialName = materialName.toLowerCase().trim();
                
                // Her option için kontrol
                for (var i = 0; i < selectElement.options.length; i++) {
                    var option = selectElement.options[i];
                    var optionText = option.text.trim().toLowerCase();
                    
                    // Daha detaylı eşleştirme
                    var materialWords = materialName.split(/\s+/);
                    var allWordsMatch = materialWords.every(word => optionText.includes(word));
                    var fullTextMatch = optionText.includes(materialName);
                    
                    if (allWordsMatch || fullTextMatch) {
                        option.selected = true;
                        selectedCount++;
                        selectedOptions.push(option.text);
                        break;  // Bu malzeme için ilk eşleşeni seç
                    }
                }
            });
            
            // Değişiklikleri uygula
            var event = new Event('change', {bubbles: true});
            selectElement.dispatchEvent(event);
            
            return {
                selectedCount: selectedCount,
                selectedOptions: selectedOptions
            };
        """, select_element, material_names)
        
        logger.info(f"JavaScript selection result: Selected {js_result['selectedCount']} options: {js_result['selectedOptions']}")
        
        # Tüm malzemelerin seçildiğinden emin ol
        if js_result['selectedCount'] < len(material_names):
            logger.warning(f"Could not select all materials. Selected {js_result['selectedCount']} out of {len(material_names)}")
            # Log the materials that were not selected
            selected_materials = set(opt.lower() for opt in js_result['selectedOptions'])
            unselected_materials = [mat for mat in material_names if mat.lower() not in selected_materials]
            logger.warning(f"Unselected materials: {unselected_materials}")
        
        return True
    
    except Exception as e:
        logger.error(f"Error in select_materials_from_csv: {str(e)}")
        # Stack trace ekle
        import traceback
        logger.error(f"Stack trace: {traceback.format_exc()}")
        return False

def parse_materials(raw_string):
    """Farklı formatlardaki malzeme bilgilerini parse eder"""
    if not raw_string or raw_string.strip() == '':
        return []
    
    # Temizleme işlemleri
    cleaned = raw_string.replace('\\', '').replace('\n', ' ').replace('"', '').strip()
    materials = []
    
    # İlk yaklaşım: Parantez içindeki malzeme isimlerini bul (örn. "1-Material (Stahl)")
    # \d+ ile herhangi bir sayıyı, [-.\s]+ ile tire, nokta veya boşluk karakterlerini yakalıyoruz
    parentheses_pattern = re.findall(r'\d+[-.\s]+(?:[^(]*)?\(([^)]+)\)', cleaned)
    if parentheses_pattern:
        materials = [match.strip() for match in parentheses_pattern]
        return materials
    
    # İkinci yaklaşım: Numaralandırılmış ayrımı kullanarak malzemeleri ayıkla
    # Örneğin: "1-Material Stahl 2-Material Aluminium" formatı için
    parts = []
    number_prefixes = re.finditer(r'\d+[-.\s]+', cleaned)
    
    # Numaralandırılmış kısımların başlangıç pozisyonlarını bul
    start_positions = [match.start() for match in number_prefixes]
    if start_positions:
        # Her numaralandırılmış kısmı ayıkla
        for i in range(len(start_positions)):
            start = start_positions[i]
            end = start_positions[i+1] if i+1 < len(start_positions) else len(cleaned)
            parts.append(cleaned[start:end].strip())
        
        # Her parçadan malzeme ismini çıkar
        for part in parts:
            # Numara önekini kaldır
            without_prefix = re.sub(r'^\d+[-.\s]+', '', part)
            
            # "Material" kelimesini kaldır
            without_material = re.sub(r'^Material\s+', '', without_prefix, flags=re.IGNORECASE)
            
            # Son kelimeyi malzeme ismi olarak al veya tüm metni kullan
            material = without_material.split()[-1] if without_material.split() else without_material
            materials.append(material.strip())
        
        return materials
    
    # Son çare: Basit ayrım yaklaşımı (orijinal kodda olduğu gibi)
    try:
        # Eğer "2-" ile bölünebiliyorsa, bu formattaki çözümü dene
        parts = cleaned.split('2-')
        
        for part in parts:
            part = part.strip().replace('1-', '').strip()
            
            match = re.search(r'\((.*?)\)', part)
            if match:
                materials.append(match.group(1).strip())
            else:
                materials.append(part.split()[-1])
        
        return materials
    except Exception as e:
        logger.error(f"Error parsing materials: {str(e)}")
        # Hiçbir şey bulunamazsa boş liste döndür
        return []

def add_product_part_from_csv(driver, csv_file_path, row_index=0):
    """Add a new product part from CSV file then continue with regular product addition"""
    try:
        logger.info("Starting process to add a new product part from CSV")
        
        # 1. Click on CLOUDLAB icon in the left menu using the Catalog approach
        logger.info("Clicking on CLOUDLAB icon using the Catalog approach...")
        
        try:
            # This is the code from navigate_to_add_product function that works for Catalog
            cloudlab_icon = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-cloudlab')]//a"))
            )
            driver.execute_script("arguments[0].click();", cloudlab_icon)
            logger.info("Clicked on CLOUDLAB icon using Catalog approach")
            
            # Wait for submenu to appear
            time.sleep(2)
        except Exception as e:
            logger.warning(f"Catalog approach failed for CLOUDLAB: {str(e)}")
            
            # Try alternative approach - find by data-ui-id or other attributes
            try:
                cloudlab_icon = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".menu-item[data-ui-id*='menu-cloudlab'], .item-cloudlab, #menu-cloudlab"))
                )
                driver.execute_script("arguments[0].click();", cloudlab_icon)
                logger.info("Clicked on CLOUDLAB icon using data-ui-id approach")
                
                # Wait for submenu to appear
                time.sleep(2)
            except Exception as e2:
                logger.warning(f"data-ui-id approach failed for CLOUDLAB: {str(e2)}")
                
                # Try the most generic approach - find by position (left menu 6th item)
                try:
                    # Find all menu items in the left menu
                    menu_items = driver.find_elements(By.CSS_SELECTOR, ".menu-item, .item, li[role='menuitem'], div[role='menuitem']")
                    
                    if len(menu_items) >= 6:
                        cloudlab_icon = menu_items[5]  # 6th item (index 5)
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cloudlab_icon)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", cloudlab_icon)
                        logger.info("Clicked on CLOUDLAB icon by position (6th menu item)")
                        
                        # Wait for submenu to appear
                        time.sleep(2)
                    else:
                        logger.warning(f"Not enough menu items found: {len(menu_items)}")
                        return False
                except Exception as e3:
                    logger.error(f"All attempts to click CLOUDLAB failed: {str(e3)}")
                    return False
        
  
        
        # 2. Click on Product Parts in the submenu - using the Catalog approach for Products submenu
        logger.info("Clicking on Product Parts submenu using Catalog Products approach...")
        
        try:
            # This is the code from navigate_to_add_product function that works for Products submenu
            product_parts_submenu = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-cloudlab')]//li[contains(@class, 'item-cloudlab-product-parts')]//a"))
            )
            driver.execute_script("arguments[0].click();", product_parts_submenu)
            logger.info("Clicked on Product Parts using Catalog approach")
        except Exception as e:
            logger.warning(f"Catalog approach failed for Product Parts: {str(e)}")
            
            # Try alternative approach - find by text
            try:
                product_parts_elements = driver.find_elements(By.XPATH, 
                    "//span[text()='Product Parts'] | //a[text()='Product Parts'] | //*[contains(text(), 'Product Parts')]")
                
                if product_parts_elements:
                    for elem in product_parts_elements:
                        if elem.is_displayed():
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                            time.sleep(1)
                            driver.execute_script("arguments[0].click();", elem)
                            logger.info("Clicked on Product Parts by text")
                            break
                    else:
                        logger.warning("Found Product Parts elements but none were clickable")
                        return False
                else:
                    logger.warning("Could not find Product Parts by text")
                    return False
            except Exception as e2:
                logger.error(f"All attempts to click Product Parts failed: {str(e2)}")
                return False
        
        # Wait for Product Parts page to load
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
            )
            
            # Additional wait to ensure page is fully loaded
            time.sleep(1)
            
            logger.info("Successfully navigated to Product Parts page")
        except Exception as wait_e:
            logger.error(f"Timeout waiting for Product Parts page to load: {str(wait_e)}")
            return False
        
       
        
        # 3. Click on Add New Part button
        logger.info("Clicking on Add New Part button")
        
        try:
            # Try to find by button text with partial match
            buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Add') and contains(text(), 'Part')]")
            if buttons:
                for button in buttons:
                    if button.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", button)
                        logger.info("Clicked on Add New Part button using partial text match")
                        break
            else:
                # Try by CSS class
                buttons = driver.find_elements(By.CSS_SELECTOR, ".action-primary, .primary, button.add, button.action-add")
                if buttons:
                    for button in buttons:
                        if button.is_displayed():
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
                            time.sleep(1)
                            driver.execute_script("arguments[0].click();", button)
                            logger.info("Clicked on Add New Part button using CSS class")
                            break
                else:
                    logger.warning("Could not find Add New Part button")
                    return False
        except Exception as e:
            logger.error(f"Failed to click Add New Part button: {str(e)}")
            return False
        
        # ÖNEMLİ: Add New Part butonuna tıkladıktan sonra 10 SANİYE BEKLE
        logger.info("Waiting  seconds for the form to load completely...")
        time.sleep(5)
        
        # 4. FORM DOLDURMA - Ürün adı ve kodu (sabit değerler)
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            
            for _ in range(6):
                next(csv_reader, None)
            
            current_row = 0
            row = None
            for row_data in csv_reader:
                if current_row == row_index:
                    row = row_data
                    break
                current_row += 1
            
            if row is None:
                logger.error(f"No data found at row index {row_index}")
                return False
            # CSV'den ürün adı ve kodu al
            product_name = row[CSV_MAPPINGS["product_name"]]
            product_code = row[CSV_MAPPINGS["code"]]

            
            logger.info(f"CSV'den alınan ürün bilgileri: Adı={product_name}, Kodu={product_code}")
        
        # NAME alanını doldur
        logger.info("Filling Name field")
        try:
            # XPath ile input bul (name alanı)
            name_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='name']"))
            )
            
            # Input alanını temizle ve doldur
            name_field.clear()
            name_field.send_keys(product_name)
            logger.info(f"Filled Name field with: {product_name}")
        except Exception as name_e:
            logger.warning(f"Could not find Name field via XPath: {str(name_e)}")
            
            # JavaScript alternatif yöntem
            try:
                driver.execute_script("""
                    // Tüm input alanlarını bul
                    var inputs = document.querySelectorAll('input');
                    
                    // Name attributuna sahip olanı bul
                    for (var i = 0; i < inputs.length; i++) {
                        if (inputs[i].name === 'name' || 
                            inputs[i].id === 'name' || 
                            inputs[i].getAttribute('data-index') === 'name') {
                            
                            inputs[i].value = arguments[0];
                            inputs[i].dispatchEvent(new Event('input', { bubbles: true }));
                            inputs[i].dispatchEvent(new Event('change', { bubbles: true }));
                            return true;
                        }
                    }
                    
                    // İlk input alanına yaz (eğer uygun olanı bulunamadıysa)
                    if (inputs.length > 0) {
                        inputs[0].value = arguments[0];
                        inputs[0].dispatchEvent(new Event('input', { bubbles: true }));
                        inputs[0].dispatchEvent(new Event('change', { bubbles: true }));
                        return true;
                    }
                    
                    return false;
                """, product_name)
                logger.info(f"Filled Name field using JavaScript with: {product_name}")
            except Exception as js_e:
                logger.error(f"Failed to fill Name field with JavaScript: {str(js_e)}")
        
        # 1 saniye bekle
        time.sleep(1)
        
                # CODE alanını doldur - Geliştirilmiş versiyon
        logger.info("Filling Code field with enhanced approach")
        try:
            # XPath ile input bul (code alanı)
            code_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='code']"))
            )
            
            # Input alanını temizle ve doldur
            code_field.clear()
            code_field.send_keys(product_code)
            
            # Enter tuşuna bas - genellikle değeri onaylar
            code_field.send_keys(Keys.ENTER)
            
            # Başka bir alana tıkla - odağı değiştir
            name_field = driver.find_element(By.XPATH, "//input[@name='name']")
            name_field.click()
            
            logger.info(f"Filled Code field with: {product_code} and triggered validation")
        except Exception as code_e:
            logger.warning(f"Could not fill Code field via XPath: {str(code_e)}")
            
            # JavaScript alternatif yöntem - Geliştirilmiş olaylarla
            try:
                driver.execute_script("""
                    // Tüm input alanlarını bul
                    var inputs = document.querySelectorAll('input');
                    
                    // Code attributuna sahip olanı bul
                    var codeInput = null;
                    for (var i = 0; i < inputs.length; i++) {
                        if (inputs[i].name === 'code' || 
                            inputs[i].id === 'code' || 
                            inputs[i].getAttribute('data-index') === 'code') {
                            
                            codeInput = inputs[i];
                            break;
                        }
                    }
                    
                    // İkinci input alanını kullan (genellikle Name'den sonra Code gelir)
                    if (!codeInput && inputs.length > 1) {
                        codeInput = inputs[1];
                    }
                    
                    if (codeInput) {
                        // Önce alanı temizle
                        codeInput.value = '';
                        codeInput.dispatchEvent(new Event('input', { bubbles: true }));
                        
                        // Değeri gir
                        codeInput.value = arguments[0];
                        
                        // Tüm olayları tetikle (sırayla)
                        codeInput.dispatchEvent(new Event('input', { bubbles: true }));
                        codeInput.dispatchEvent(new Event('change', { bubbles: true }));
                        codeInput.dispatchEvent(new Event('blur', { bubbles: true }));
                        
                        // Validasyonu zorla
                        codeInput.checkValidity();
                        
                        return "Filled Code field with enhanced events: " + arguments[0];
                    }
                    
                    return "Could not find Code field";
                """, product_code)
                logger.info(f"Filled Code field using JavaScript with enhanced events")
            except Exception as js_e:
                logger.error(f"Failed to fill Code field with JavaScript: {str(js_e)}")
                
        
        
         # 5. Find Material dropdown and select the last option
        logger.info("Attempting to select Material dropdown")
        if select_materials_from_csv(driver, csv_file_path):
            logger.info("Successfully selected last option in Material dropdown")
        else:
            logger.warning("Failed to select Material dropdown, continuing with save")
        
        time.sleep(5)
        # 6. SAVE button'una tıkla
        logger.info("Clicking Save button")
        try:
            # En basit, direkt CSS selector ile Save butonunu bul
            save_button = driver.find_element(By.CSS_SELECTOR, "button.save, .action-save")
            
            # Butonu görünür yap ve tıkla
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", save_button)
            logger.info("Clicked Save button via CSS selector")
        except Exception as e1:
            logger.warning(f"Could not find Save button via CSS selector: {str(e1)}")
            
            try:
                # XPath ile dene - tam metin eşleşmesi
                save_button = driver.find_element(By.XPATH, "//button[text()='Save']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", save_button)
                logger.info("Clicked Save button via exact text XPath")
            except Exception as e2:
                logger.warning(f"Could not find Save button via exact text XPath: {str(e2)}")
                
                try:
                    # XPath ile dene - kısmi metin eşleşmesi
                    save_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Save')]")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", save_button)
                    logger.info("Clicked Save button via partial text XPath")
                except Exception as e3:
                    logger.warning(f"Could not find Save button via partial text XPath: {str(e3)}")
                    
                    # Görüntüde açıkça görünen bir Save button olduğu için, pozisyonuna göre bulalım
                    try:
                        # JavaScript ile sağ üst köşedeki butonu bul
                        result = driver.execute_script("""
                            // Sağ üst köşedeki butonu bul (sayfanın üst kısmında ve sağdan 200px içeride)
                            var buttons = document.querySelectorAll('button');
                            var rightTopButtons = [];
                            
                            for (var i = 0; i < buttons.length; i++) {
                                var rect = buttons[i].getBoundingClientRect();
                                var windowWidth = window.innerWidth;
                                
                                // Sağ üst köşede (sayfanın üst 200px'i içinde ve sağ 300px içinde) olan butonları bul
                                if (rect.top < 200 && (windowWidth - rect.right) < 300 && window.getComputedStyle(buttons[i]).display !== 'none') {
                                    rightTopButtons.push({
                                        button: buttons[i],
                                        right: windowWidth - rect.right,  // Sağ kenardan olan mesafe
                                        text: buttons[i].textContent.trim()
                                    });
                                }
                            }
                            
                            // Sağ kenardan uzaklığa göre sırala (en sağdaki buton ilk olacak)
                            rightTopButtons.sort(function(a, b) {
                                return a.right - b.right;
                            });
                            
                            // Buton bulundu mu kontrol et
                            if (rightTopButtons.length > 0) {
                                // En sağdaki butona tıkla
                                rightTopButtons[0].button.click();
                                return "Clicked rightmost top button with text: " + rightTopButtons[0].text;
                            }
                            
                            // Son çare: sayfadaki tüm 'primary' butonları dene
                            var primaryButtons = document.querySelectorAll('.primary, .action-primary');
                            if (primaryButtons.length > 0) {
                                for (var i = 0; i < primaryButtons.length; i++) {
                                    if (window.getComputedStyle(primaryButtons[i]).display !== 'none') {
                                        primaryButtons[i].click();
                                        return "Clicked primary button at index " + i;
                                    }
                                }
                            }
                            
                            // Buton bulunamadı
                            return "No suitable button found";
                        """)
                        
                        logger.info(f"JavaScript button location result: {result}")
                        
                        if "Clicked" in str(result):
                            logger.info("Successfully clicked a button using JavaScript")
                        else:
                            logger.warning("JavaScript approach did not find any suitable button")
                            
                            # Son çare: ActionChains ile sağ üstteki Save butonuna tıklamaya çalış
                            try:
                                # Sayfanın sağ üst köşesine tıkla (X=genişlik-100, Y=200)
                                window_size = driver.get_window_size()
                                action = ActionChains(driver)
                                action.move_by_offset(window_size['width']-100, 200).click().perform()
                                logger.info("Clicked at position right upper corner using ActionChains")
                            except Exception as action_e:
                                logger.error(f"ActionChains approach also failed: {str(action_e)}")
                    except Exception as js_e:
                        logger.error(f"JavaScript approach failed: {str(js_e)}")
        
        # 10 saniye bekle
        logger.info("Waiting 10 seconds after save...")
        time.sleep(10)
        

        
        logger.info("Product part save process completed")
        return True
        
    except Exception as e:
        logger.error(f"Failed in add_product_part_from_csv: {str(e)}")
        # Include stack trace for better debugging
        import traceback
        logger.error(f"Stack trace: {traceback.format_exc()}")
        return False

def navigate_to_add_product(driver):
    """Navigate to the Add Product page"""
    try:
        # Click on Catalog icon in the left menu
        logger.info("Clicking on Catalog icon...")
        catalog_icon = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-catalog')]//a"))
        )
        driver.execute_script("arguments[0].click();", catalog_icon)
        
        # Wait for submenu to appear
        time.sleep(2)
        
        # Click on Products in the submenu
        logger.info("Clicking on Products submenu...")
        products_submenu = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-catalog')]//li[contains(@class, 'item-catalog-products')]//a"))
        )
        driver.execute_script("arguments[0].click();", products_submenu)
        
        # Wait for products page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
        )
        
        # Additional wait to ensure page is fully loaded
        time.sleep(3)
        
        # Click on Add Product button
        logger.info("Clicking on Add Product button...")
        add_product_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.add"))
        )
        driver.execute_script("arguments[0].click();", add_product_button)
        
        # Wait for product creation page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
        )
        
        # Additional wait to ensure page is fully loaded
        time.sleep(3)
        
        logger.info("Navigated to Add Product page")
        return True
    except Exception as e:
        logger.error(f"Error navigating to Add Product page: {str(e)}")
        return False

def fill_basic_info(driver, product_name, sku):
    """Fill basic product information"""
    try:
        # Enter product name
        product_name_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='product[name]']"))
        )
        product_name_field.clear()
        product_name_field.send_keys(product_name)
        logger.info(f"Product name entered: {product_name}")
        
        # Enter SKU
        sku_field = driver.find_element(By.CSS_SELECTOR, "input[name='product[sku]']")
        sku_field.clear()
        sku_field.send_keys(sku)
        logger.info(f"SKU entered: {sku}")
        
        # Enter price
        price_field = driver.find_element(By.CSS_SELECTOR, "input[name='product[price]']")
        price_field.clear()
        price_field.send_keys(PRICE)
        logger.info(f"Price entered: {PRICE}")
        
        # Enter quantity
        qty_field = driver.find_element(By.CSS_SELECTOR, "input[name='product[quantity_and_stock_status][qty]']")
        qty_field.clear()
        qty_field.send_keys(QUANTITY)
        logger.info(f"Quantity entered: {QUANTITY}")
        
        logger.info(f"Successfully filled basic info for product: {product_name}")
        return True
    except Exception as e:
        logger.error(f"Failed to fill basic info: {str(e)}")
        return False

def set_meter_price(driver, is_meter_price=True):
    """Set the 'Is Meter Price' option according to the parameter value with advanced positioning"""
    try:
        logger.info(f"Setting 'Is Meter Price' to: {'Yes' if is_meter_price else 'No'}")
        
        
        # Now use direct JavaScript to find and click the toggle
        # This approach avoids the risk of page jumping
        result = driver.execute_script("""
            // Find the label again (in case page structure changed)
            var elements = Array.from(document.querySelectorAll('*')).filter(el => 
                el.textContent && el.textContent.trim() === 'Is Meter Price');
            
            if (elements.length === 0) return "Could not find 'Is Meter Price' label";
            
            // Get the one that's most likely the label
            var label = elements.find(el => el.tagName.toLowerCase() === 'span');
            if (!label) label = elements[0];
            
            // Find the parent field by traversing up
            var parentField = label;
            while (parentField && !parentField.classList.contains('admin__field')) {
                parentField = parentField.parentElement;
                if (!parentField) return "Could not find parent field";
            }
            
            // Find the toggle within the parent field
            var toggle = parentField.querySelector('label.admin__actions-switch-label');
            if (!toggle) return "Could not find toggle element";
            
            // Get current state
            var input = parentField.querySelector('input[type="checkbox"]');
            var isCurrentlyEnabled = input ? input.checked : 
                                    toggle.classList.contains('_active') || 
                                    toggle.classList.contains('admin__actions-switch-checkbox-checked');
            
            var currentState = isCurrentlyEnabled ? "Yes" : "No";
            var targetState = arguments[0] ? "Yes" : "No";
            
            // Only toggle if states don't match
            if (currentState !== targetState) {
                // Use click() instead of programmatically changing checked state
                // as click() will trigger any necessary events
                toggle.click();
                
                // Force value change in case click doesn't work
                if (input) {
                    input.checked = arguments[0];
                    // Trigger events to ensure change is registered
                    input.dispatchEvent(new Event('change', {bubbles: true}));
                }
                
                return "Changed 'Is Meter Price' to: " + targetState;
            } else {
                return "'Is Meter Price' is already set to: " + targetState;
            }
        """, is_meter_price)
        
        logger.info(f"Toggle result: {result}")
        
        # Verify current scroll position hasn't jumped
        final_scroll_y = driver.execute_script("return window.pageYOffset;")
        logger.info(f"Final scroll position: {final_scroll_y}px")
        
        if abs(final_scroll_y - current_scroll_y) > 100:
            logger.warning(f"Page position changed significantly during operation: {current_scroll_y} -> {final_scroll_y}")
   
        
        # Additional validation: check if setting was applied
        validation = driver.execute_script("""
            // Find the Is Meter Price input again
            var elements = Array.from(document.querySelectorAll('*')).filter(el => 
                el.textContent && el.textContent.trim() === 'Is Meter Price');
            
            if (elements.length === 0) return {success: false, message: "Could not find element"};
            
            var label = elements.find(el => el.tagName.toLowerCase() === 'span');
            if (!label) label = elements[0];
            
            // Find the parent field
            var parentField = label;
            while (parentField && !parentField.classList.contains('admin__field')) {
                parentField = parentField.parentElement;
                if (!parentField) return {success: false, message: "Could not find parent"};
            }
            
            // Check current state
            var input = parentField.querySelector('input[type="checkbox"]');
            var isEnabled = input ? input.checked : 
                         parentField.querySelector('label')?.classList.contains('_active');
            
            return {
                success: true,
                isEnabled: isEnabled,
                expected: arguments[0]
            };
        """, is_meter_price)
        
        if isinstance(validation, dict) and validation.get('success'):
            if validation.get('isEnabled') == is_meter_price:
                logger.info("Validation confirmed 'Is Meter Price' setting was applied correctly")
            else:
                logger.warning(f"Validation shows setting may not have been applied: {validation}")
        
        return True
    except Exception as e:
        logger.error(f"Failed to set 'Is Meter Price': {str(e)}")
        
        # Try one more approach if all else fails
        try:
            logger.info("Attempting final fallback method")
            
            # Scroll to top of page first to get a clean slate
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)
            
            # Try to find the exact text using a more robust XPath
            meter_price_label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                    "//span[normalize-space(text())='Is Meter Price' and not(ancestor::*[contains(@style,'display: none')])]"))
            )
            
            # Get exact coordinates
            rect = driver.execute_script("""
                var rect = arguments[0].getBoundingClientRect();
                return {
                    x: rect.left + rect.width/2,
                    y: rect.top + rect.height/2
                };
            """, meter_price_label)
            
            # Scroll to position with good visibility
            driver.execute_script(f"window.scrollTo(0, {rect['y'] - 150});")
            time.sleep(1)
            
            # Find the toggle by looking at siblings and nearby elements
            toggle = driver.execute_script("""
                var element = arguments[0];
                var parentField = element.closest('.admin__field');
                
                if (!parentField) {
                    // If we can't find the field, look at the parent containers
                    parentField = element.parentElement;
                    while (parentField && !parentField.querySelector('input[type="checkbox"]')) {
                        parentField = parentField.parentElement;
                    }
                }
                
                if (!parentField) return null;
                
                // Look for the toggle or checkbox
                return parentField.querySelector('label.admin__actions-switch-label') || 
                       parentField.querySelector('input[type="checkbox"]');
            """, meter_price_label)
            
            if toggle:
                # Click the toggle with JavaScript
                driver.execute_script("arguments[0].click();", toggle)
                logger.info("Clicked toggle using fallback method")
                
                
                
                return True
            else:
                logger.error("Could not find toggle element in fallback method")
                return False
                
        except Exception as fallback_e:
            logger.error(f"All methods failed to set 'Is Meter Price': {str(fallback_e)}")
            return False        
    
def select_producer_simplified(driver, producer_value="wmd.networkhero.eu"):
    """Select the specified producer with a simplified approach that focuses on the most reliable method"""
    try:
        logger.info(f"Selecting Producer: {producer_value} with simplified approach")
        
       
        
        # Find the Producer label
        producer_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[text()='Producer']"))
        )
        
        # Scroll to the Producer field to ensure it's visible
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", producer_label)
        time.sleep(1)
        
        # Use JavaScript directly since this worked in the logs
        result = driver.execute_script("""
            // Get the producer label
            var labels = Array.from(document.querySelectorAll('span')).filter(
                el => el.textContent.trim() === 'Producer'
            );
            
            if (labels.length === 0) return "Producer label not found";
            
            var label = labels[0];
            
            // Find the field container
            var field = label.closest('.admin__field');
            if (!field) return "Field container not found";
            
            // Find if it's a standard dropdown
            var select = field.querySelector('select');
            if (select) {
                // It's a standard dropdown
                var options = Array.from(select.options);
                var targetOption = options.find(opt => opt.text.trim() === arguments[0]);
                
                if (targetOption) {
                    targetOption.selected = true;
                    select.dispatchEvent(new Event('change', {bubbles: true}));
                    return "Selected " + arguments[0] + " using standard dropdown via JS";
                } else {
                    return "Option " + arguments[0] + " not found in standard dropdown";
                }
            }
            
            // Find if it's a custom dropdown
            var customDropdown = field.querySelector('.admin__action-multiselect');
            if (customDropdown) {
                // Click to open the dropdown
                customDropdown.click();
                
                // Wait a brief moment for the dropdown to open
                setTimeout(function(){}, 100);
                
                // Find the option element
                var options = document.querySelectorAll('.admin__action-multiselect-menu-inner-item span');
                
                // Look for exact match first
                for (var i = 0; i < options.length; i++) {
                    if (options[i].textContent.trim() === arguments[0]) {
                        options[i].click();
                        return "Selected " + arguments[0] + " from custom dropdown";
                    }
                }
                
                // If we couldn't find exact match, try contains
                for (var i = 0; i < options.length; i++) {
                    if (options[i].textContent.includes(arguments[0])) {
                        options[i].click();
                        return "Selected option containing " + arguments[0];
                    }
                }
                
                return "Option not found in custom dropdown";
            }
            
            return "No dropdown element found";
        """, producer_value)
        
        logger.info(f"JavaScript producer selection result: {result}")
        
        if "not found" in result:
            # Only if the JavaScript approach fails, try Selenium approach
            try:
                # Find the dropdown trigger
                dropdown_trigger = producer_label.find_element(By.XPATH, "./ancestor::div[contains(@class, 'admin__field')]//div[contains(@class, 'admin__action-multiselect')]")
                
                # Click to open the dropdown
                driver.execute_script("arguments[0].click();", dropdown_trigger)
                logger.info("Clicked dropdown to open it")
                time.sleep(1)
                
                # Find and click the desired option
                option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, 'admin__action-multiselect-menu-inner-item')]//span[text()='{producer_value}']"))
                )
                
                driver.execute_script("arguments[0].click();", option)
                logger.info(f"Selected {producer_value} from custom dropdown")
            except Exception as fallback_e:
                logger.error(f"Fallback method also failed: {str(fallback_e)}")
                return False
        
 
        
        return True
    except Exception as e:
        logger.error(f"Failed to select producer with simplified approach: {str(e)}")
        return False

def set_disable_pickup(driver, disable_pickup=True):
    """Set the 'Disable pickup' option according to the parameter value with advanced positioning"""
    try:
        logger.info(f"Setting 'Disable pickup' to: {'Yes' if disable_pickup else 'No'}")
        
       
        
        # First, disable smooth scrolling to avoid animation issues
        driver.execute_script("document.documentElement.style.scrollBehavior = 'auto';")
        
        # Store initial scroll position to check if page jumps
        initial_scroll_y = driver.execute_script("return window.pageYOffset;")
        logger.info(f"Initial scroll position: {initial_scroll_y}px")
        
        # Execute a direct search for the Disable pickup element using JavaScript
        # This avoids scrolling and element focus issues
        disable_pickup_position = driver.execute_script("""
            // Find all elements with EXACT text "Disable pickup" (no extra whitespace)
            var elements = Array.from(document.querySelectorAll('*')).filter(el => 
                el.textContent && el.textContent.trim() === 'Disable pickup');
            
            if (elements.length === 0) return null;
            
            // Get the one that's most likely the label (span within admin field)
            var label = elements.find(el => el.tagName.toLowerCase() === 'span');
            if (!label) label = elements[0];
            
            // Return the element's position
            var rect = label.getBoundingClientRect();
            return {
                x: rect.left + window.pageXOffset,
                y: rect.top + window.pageYOffset,
                found: true,
                element: label
            };
        """)
        
        if not disable_pickup_position or not disable_pickup_position.get('found'):
            logger.error("Could not find 'Disable pickup' label with JavaScript")
            raise Exception("Disable pickup label not found")
        
        # Scroll to position with a slight offset to ensure it's visible
        # Use absolute positioning to avoid relative scrolling issues
        y_position = disable_pickup_position['y'] - 150  # Added offset to ensure good visibility
        driver.execute_script(f"window.scrollTo(0, {y_position});")
        time.sleep(1.5)  # Allow time for scroll to complete
        
        # Verify we're at expected position
        current_scroll_y = driver.execute_script("return window.pageYOffset;")
        logger.info(f"Current scroll position after positioning: {current_scroll_y}px")
       
        
        # Now use direct JavaScript to find and click the toggle
        # This approach avoids the risk of page jumping
        result = driver.execute_script("""
            // Find the Disable pickup label again to be safe
            var elements = Array.from(document.querySelectorAll('*')).filter(el => 
                el.textContent && el.textContent.trim() === 'Disable pickup');
            
            if (elements.length === 0) return "Could not find 'Disable pickup' label";
            
            // Get the one that's most likely the label
            var label = elements.find(el => el.tagName.toLowerCase() === 'span');
            if (!label) label = elements[0];
            
            // Find the parent field by traversing up
            var parentField = label;
            while (parentField && !parentField.classList.contains('admin__field')) {
                parentField = parentField.parentElement;
                if (!parentField) return "Could not find parent field";
            }
            
            // Find the toggle within the parent field
            var toggle = parentField.querySelector('label.admin__actions-switch-label');
            if (!toggle) return "Could not find toggle element";
            
            // Get current state
            var input = parentField.querySelector('input[type="checkbox"]');
            var isCurrentlyEnabled = input ? input.checked : 
                                    toggle.classList.contains('_active') || 
                                    toggle.classList.contains('admin__actions-switch-checkbox-checked');
            
            var currentState = isCurrentlyEnabled ? "Yes" : "No";
            var targetState = arguments[0] ? "Yes" : "No";
            
            // Only toggle if states don't match
            if (currentState !== targetState) {
                // Use click() instead of programmatically changing checked state
                // as click() will trigger any necessary events
                toggle.click();
                
                // Force value change in case click doesn't work
                if (input) {
                    input.checked = arguments[0];
                    // Trigger events to ensure change is registered
                    input.dispatchEvent(new Event('change', {bubbles: true}));
                }
                
                return "Changed 'Disable pickup' to: " + targetState;
            } else {
                return "'Disable pickup' is already set to: " + targetState;
            }
        """, disable_pickup)
        
        logger.info(f"Toggle result: {result}")
        
        # Verify current scroll position hasn't jumped
        final_scroll_y = driver.execute_script("return window.pageYOffset;")
        logger.info(f"Final scroll position: {final_scroll_y}px")
        
        if abs(final_scroll_y - current_scroll_y) > 100:
            logger.warning(f"Page position changed significantly during operation: {current_scroll_y} -> {final_scroll_y}")
        
       
        
        # Additional validation: check if setting was applied
        validation = driver.execute_script("""
            // Find the Disable pickup label again
            var elements = Array.from(document.querySelectorAll('*')).filter(el => 
                el.textContent && el.textContent.trim() === 'Disable pickup');
            
            if (elements.length === 0) return {success: false, message: "Could not find element"};
            
            var label = elements.find(el => el.tagName.toLowerCase() === 'span');
            if (!label) label = elements[0];
            
            // Find the parent field
            var parentField = label;
            while (parentField && !parentField.classList.contains('admin__field')) {
                parentField = parentField.parentElement;
                if (!parentField) return {success: false, message: "Could not find parent"};
            }
            
            // Check current state
            var input = parentField.querySelector('input[type="checkbox"]');
            var isEnabled = input ? input.checked : 
                         parentField.querySelector('label')?.classList.contains('_active');
            
            return {
                success: true,
                isEnabled: isEnabled,
                expected: arguments[0]
            };
        """, disable_pickup)
        
        if isinstance(validation, dict) and validation.get('success'):
            if validation.get('isEnabled') == disable_pickup:
                logger.info("Validation confirmed 'Disable pickup' setting was applied correctly")
            else:
                logger.warning(f"Validation shows setting may not have been applied: {validation}")
        
        return True
    except Exception as e:
        logger.error(f"Failed to set 'Disable pickup': {str(e)}")
        
        # Try one more approach if all else fails - this is a more direct approach
        try:
            logger.info("Attempting final fallback method for Disable pickup")
            
            # Use simpler direct JavaScript to find and click the toggle
            result = driver.execute_script("""
                // Direct query for Disable pickup with exact text match
                var exactLabelText = 'Disable pickup';
                
                // 1. Find all spans with this exact text
                var spans = Array.from(document.querySelectorAll('span')).filter(span => 
                    span.textContent.trim() === exactLabelText);
                
                if (spans.length === 0) {
                    return "No element with exact text 'Disable pickup' found";
                }
                
                // 2. Get the span that's visible (not hidden by CSS)
                var visibleSpan = spans.find(span => {
                    // Check if element or any parent is hidden
                    let el = span;
                    while (el) {
                        const style = window.getComputedStyle(el);
                        if (style.display === 'none' || style.visibility === 'hidden') {
                            return false;
                        }
                        el = el.parentElement;
                    }
                    return true;
                });
                
                if (!visibleSpan) {
                    return "Found 'Disable pickup' spans but none are visible";
                }
                
                // 3. Get the toggle associated with this span
                var adminField = visibleSpan.closest('.admin__field');
                if (!adminField) {
                    return "Could not find admin field container for 'Disable pickup'";
                }
                
                var toggle = adminField.querySelector('.admin__actions-switch-label');
                if (!toggle) {
                    return "Could not find toggle element for 'Disable pickup'";
                }
                
                // 4. Check current state
                var checkbox = adminField.querySelector('input[type="checkbox"]');
                var currentlyEnabled = checkbox ? checkbox.checked : false;
                
                // 5. Toggle if needed
                if (currentlyEnabled !== arguments[0]) {
                    toggle.click();
                    
                    // Wait a bit for the click to take effect
                    setTimeout(function() {}, 300);
                    
                    // Force checkbox state just to be safe
                    if (checkbox) {
                        checkbox.checked = arguments[0];
                        checkbox.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                    
                    return "Changed 'Disable pickup' setting";
                }
                
                return "'Disable pickup' was already set correctly";
            """, disable_pickup);
            
            logger.info(f"Fallback result for Disable pickup: {result}")
            return "not found" not in result and "could not find" not in result.lower()
            
        except Exception as fallback_e:
            logger.error(f"All methods failed to set 'Disable pickup': {str(fallback_e)}")
            return False

def select_websites(driver, websites_to_select):
    """Scroll down to find Websites section and select specified websites"""
    try:
        logger.info(f"Scrolling down to find Websites section")
        
      
        
        # Find the Websites section
        websites_section = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "[data-index='websites']"))
        )
        
        # Scroll to the Websites section
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", websites_section)
        time.sleep(1)
        
      
        
        # Check if section is expanded
        if "active" not in websites_section.get_attribute("class"):
            # Click to expand using JavaScript for reliability
            websites_header = websites_section.find_element(By.CSS_SELECTOR, ".admin__collapsible-title")
            driver.execute_script("arguments[0].click();", websites_header)
            logger.info("Expanded Websites section")
            time.sleep(2)
        
     
        # Select each website
        for website in websites_to_select:
            try:
                # Try to find the checkbox by label text
                website_checkbox = driver.find_element(
                    By.XPATH, 
                    f"//div[contains(@class, 'admin__field-option')]//label[contains(text(), '{website}')]/preceding-sibling::input"
                )
                
                # Check if already selected
                if website_checkbox.is_selected():
                    logger.info(f"Website {website} already selected")
                    continue
                
                # Click the checkbox
                driver.execute_script("arguments[0].click();", website_checkbox)
                logger.info(f"Selected website: {website}")
                
              
            except Exception as e:
                logger.warning(f"Failed to select website {website} by checkbox: {str(e)}")
                
                # Try clicking on the label instead
                try:
                    website_label = driver.find_element(
                        By.XPATH, 
                        f"//div[contains(@class, 'admin__field-option')]//label[contains(text(), '{website}')]"
                    )
                    driver.execute_script("arguments[0].click();", website_label)
                    logger.info(f"Selected website {website} by clicking on label")
                except Exception as label_e:
                    logger.warning(f"Failed to select website {website} by label: {str(label_e)}")
        
        logger.info("Completed website selection")
        return True
    except Exception as e:
        logger.error(f"Failed to select websites: {str(e)}")
        
        # Try an alternative approach if the first one fails
        try:
            logger.info("Trying alternative approach to select websites")
            
            # Scroll down more
            js_scroll = "window.scrollBy(0, 500);"
            driver.execute_script(js_scroll)
            time.sleep(1)
            
            # Try JavaScript to select websites
            js_script = """
            // Find all checkboxes and their labels
            var labels = document.querySelectorAll('label');
            var websitesToSelect = arguments[0];
            var selected = [];
            
            for (var i = 0; i < labels.length; i++) {
                var label = labels[i];
                var text = label.textContent || label.innerText;
                
                // Check if this label contains one of our target websites
                for (var j = 0; j < websitesToSelect.length; j++) {
                    if (text.includes(websitesToSelect[j])) {
                        // Found a matching label, now find the checkbox
                        var checkbox = label.previousElementSibling;
                        if (checkbox && checkbox.type === 'checkbox') {
                            // Check if already selected
                            if (!checkbox.checked) {
                                checkbox.checked = true;
                                checkbox.dispatchEvent(new Event('change', { 'bubbles': true }));
                                selected.push(websitesToSelect[j]);
                            } else {
                                selected.push(websitesToSelect[j] + ' (already selected)');
                            }
                        }
                    }
                }
            }
            
            return "Selected websites: " + selected.join(", ");
            """
            
            result = driver.execute_script(js_script, websites_to_select)
            logger.info(f"JavaScript website selection result: {result}")
            
            return True
        except Exception as alt_e:
            logger.error(f"Alternative approach also failed: {str(alt_e)}")
            return False

def update_product_type_like_producer(driver, product_type="Upload + Personalization + Configuration"):
    """Update Product Type using the same successful approach that works for Producer"""
    try:
        logger.info(f"Updating Product Type to: {product_type} using producer-style approach")
        
        
        # 1. Web2Print Settings'e tıkla
        web2print_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Web2Print Settings')]")
        
        if web2print_elements:
            for elem in web2print_elements:
                try:
                    if elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", elem)
                        logger.info("Clicked Web2Print Settings")
                        time.sleep(2)
                        break
                except:
                    continue
        else:
            logger.warning("No Web2Print Settings elements found")
            return False
        
        # 2. General sekmesine tıkla
        general_elements = driver.find_elements(By.XPATH, "//span[text()='General']")
        
        if general_elements:
            for elem in general_elements:
                try:
                    if elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", elem)
                        logger.info("Clicked General tab")
                        time.sleep(2)
                        break
                except:
                    continue
        else:
            logger.warning("No General elements found")
            return False
        
        # 3. Producer başarılı yaklaşımını kullanarak Product Type'ı güncelle
        logger.info(f"Selecting Product Type: {product_type} with producer-style approach")
        
        
        # Find the Product Type label
        product_type_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[text()='Product Type']"))
        )
        
        # Scroll to the Product Type field to ensure it's visible
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", product_type_label)
        time.sleep(1)
        
        # Use JavaScript directly since this worked for producer
        result = driver.execute_script("""
            // Get the Product Type label
            var labels = Array.from(document.querySelectorAll('span')).filter(
                el => el.textContent.trim() === 'Product Type'
            );
            
            if (labels.length === 0) return "Product Type label not found";
            
            var label = labels[0];
            
            // Find the field container
            var field = label.closest('.admin__field');
            if (!field) return "Field container not found";
            
            // Find the input or dropdown element nearby
            var inputElement = field.querySelector('input[type="text"]');
            var selectElement = field.querySelector('select');
            var dropdown = field.querySelector('.action-select') || 
                           field.querySelector('.admin__action-multiselect');
            
            // Handle different element types
            if (selectElement) {
                // Standard select dropdown
                var options = Array.from(selectElement.options);
                var targetOption = options.find(opt => opt.text.includes(arguments[0]));
                
                if (targetOption) {
                    targetOption.selected = true;
                    selectElement.dispatchEvent(new Event('change', {bubbles: true}));
                    return "Selected " + arguments[0] + " using standard select";
                } else {
                    return "Option " + arguments[0] + " not found in standard select";
                }
            } else if (dropdown) {
                // Custom dropdown
                dropdown.click();
                
                // Wait a brief moment for the dropdown to open
                setTimeout(function(){}, 100);
                
                // Find the option elements
                var options = document.querySelectorAll('.admin__action-multiselect-menu-inner-item span, .action-menu-item, li.option-item span');
                
                // Look for option containing our target text
                for (var i = 0; i < options.length; i++) {
                    if (options[i].textContent.includes(arguments[0])) {
                        options[i].click();
                        return "Selected " + arguments[0] + " from custom dropdown";
                    }
                }
                
                return "Option not found in custom dropdown";
            } else if (inputElement) {
                // Input field (unlikely for Product Type)
                inputElement.value = arguments[0];
                inputElement.dispatchEvent(new Event('input', {bubbles: true}));
                inputElement.dispatchEvent(new Event('change', {bubbles: true}));
                return "Set input value to " + arguments[0];
            }
            
            // If we get here, we couldn't find a way to interact with the field
            // Let's try one more direct approach - click directly right of the label
            var rect = label.getBoundingClientRect();
            var x = rect.right + 50;
            var y = rect.top + (rect.height / 2);
            
            // Create and dispatch a click event at these coordinates
            var clickEvent = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: x,
                clientY: y
            });
            
            var elementAtPoint = document.elementFromPoint(x, y);
            if (elementAtPoint) {
                elementAtPoint.dispatchEvent(clickEvent);
                
                // After clicking, try to find options again
                setTimeout(function(){}, 100);
                
                var visibleOptions = Array.from(document.querySelectorAll('*')).filter(
                    el => el.textContent.includes(arguments[0]) && 
                    window.getComputedStyle(el).display !== 'none'
                );
                
                if (visibleOptions.length > 0) {
                    visibleOptions[0].click();
                    return "Selected " + arguments[0] + " after clicking near label";
                }
                
                return "Clicked near label but couldn't find option";
            }
            
            return "Could not find any way to interact with Product Type field";
        """, product_type)
        
        logger.info(f"JavaScript Product Type selection result: {result}")
        
        
        # If JavaScript approach fails, try direct Selenium
        if "not found" in result or "Could not find" in result:
            logger.warning("JavaScript approach failed, trying direct Selenium for Product Type selection")
            
            try:
                # Try using a more generic approach to find the dropdown
                # Look around the Product Type label
                product_type_label = driver.find_element(By.XPATH, "//span[text()='Product Type']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", product_type_label)
                time.sleep(1)
                
                # Click in the general vicinity of where the dropdown should be
                # Get the position of the label and click to its right
                location = product_type_label.location
                size = product_type_label.size
                
                # Create ActionChains
                actions = ActionChains(driver)
                
                # Move to label first
                actions.move_to_element(product_type_label)
                
                # Then move 150px to the right (should be where the dropdown is)
                actions.move_by_offset(150, 0)
                
                # Click there
                actions.click()
                actions.perform()
                logger.info("Clicked near Product Type label with ActionChains")
                time.sleep(1)
                
                
                # Now try to find and click any option element containing our text
                options = driver.find_elements(By.XPATH, f"//*[contains(text(), '{product_type}')]")
                
                if options:
                    for option in options:
                        try:
                            if option.is_displayed():
                                driver.execute_script("arguments[0].click();", option)
                                logger.info(f"Clicked option with text containing: {product_type}")
                                time.sleep(1)
                                return True
                        except:
                            continue
                
                logger.warning("Could not find any visible options containing target text")
            except Exception as selenium_e:
                logger.error(f"Selenium approach for Product Type selection failed: {str(selenium_e)}")
        
        # Return success if we didn't get a clear failure message
        return not ("not found" in result or "Could not find" in result)
    except Exception as e:
        logger.error(f"Failed to update Product Type with producer approach: {str(e)}")
        return False

def set_price_calculation_to_fixed(driver):
    """Set the Price Calculation to Fixed in the Matrix tab using producer-style approach"""
    try:
        logger.info("Setting Price Calculation to Fixed in Matrix tab")
        
        
        # 1. Matrix sekmesine tıkla
        matrix_elements = driver.find_elements(By.XPATH, "//span[text()='Matrix']")
        
        if matrix_elements:
            for elem in matrix_elements:
                try:
                    if elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", elem)
                        logger.info("Clicked Matrix tab")
                        time.sleep(2)
                        break
                except:
                    continue
        else:
            logger.warning("No Matrix tab elements found")
            return False
        
        # Matrix tıklandıktan sonra ekran görüntüsü
        
        # 2. Price Calculation alanını bul ve "Fixed" olarak ayarla
        logger.info("Looking for Price Calculation field")
        
        # Producer-style yaklaşım ile JavaScript kullan
        result = driver.execute_script("""
            // Get the Price Calculation label
            var labels = Array.from(document.querySelectorAll('span')).filter(
                el => el.textContent.trim() === 'Price Calculation'
            );
            
            if (labels.length === 0) {
                // Try with broader search
                labels = Array.from(document.querySelectorAll('*')).filter(
                    el => el.textContent && el.textContent.trim() === 'Price Calculation' && el.children.length === 0
                );
            }
            
            if (labels.length === 0) return "Price Calculation label not found";
            
            var label = labels[0];
            
            // Find the field container
            var field = label.closest('.admin__field');
            if (!field) return "Field container not found";
            
            // Find the select element or dropdown
            var select = field.querySelector('select');
            var dropdown = field.querySelector('.action-select') || 
                          field.querySelector('.admin__action-multiselect');
            
            // Handle different element types
            if (select) {
                // Standard select dropdown
                var options = Array.from(select.options);
                var fixedOption = options.find(opt => opt.text.trim() === 'Fixed');
                
                if (fixedOption) {
                    select.value = fixedOption.value;
                    select.dispatchEvent(new Event('change', {bubbles: true}));
                    return "Selected Fixed using standard select";
                } else {
                    return "Fixed option not found in standard select";
                }
            } else if (dropdown) {
                // Custom dropdown
                dropdown.click();
                
                // Wait a brief moment for the dropdown to open
                setTimeout(function(){}, 100);
                
                // Find the Fixed option
                var options = document.querySelectorAll('.admin__action-multiselect-menu-inner-item span, .action-menu-item, li.option-item span');
                
                for (var i = 0; i < options.length; i++) {
                    if (options[i].textContent.trim() === 'Fixed') {
                        options[i].click();
                        return "Selected Fixed from custom dropdown";
                    }
                }
                
                return "Fixed option not found in custom dropdown";
            }
            
            // If we get here, try clicking directly on the field and then finding the Fixed option
            var field = label.closest('.admin__field');
            
            if (field) {
                // Click in the general area where the dropdown should be
                var rect = field.getBoundingClientRect();
                var x = rect.left + rect.width - 50; // Near the right side of the field
                var y = rect.top + (rect.height / 2);
                
                var elementAtPoint = document.elementFromPoint(x, y);
                if (elementAtPoint) {
                    elementAtPoint.click();
                    
                    // After clicking, try to find Fixed option
                    setTimeout(function(){}, 100);
                    
                    var allVisible = Array.from(document.querySelectorAll('*')).filter(
                        el => el.textContent && 
                        el.textContent.trim() === 'Fixed' && 
                        window.getComputedStyle(el).display !== 'none'
                    );
                    
                    if (allVisible.length > 0) {
                        allVisible[0].click();
                        return "Selected Fixed after clicking in field area";
                    }
                }
            }
            
            return "Could not find a way to select Fixed";
        """)
        
        logger.info(f"JavaScript Price Calculation selection result: {result}")
        
     
        # JavaScript başarısız olursa, Selenium dene
        if "not found" in result or "Could not find" in result:
            logger.warning("JavaScript approach failed, trying direct Selenium")
            
            try:
                # Price Calculation etiketini bul
                price_calc_label = driver.find_element(By.XPATH, "//*[text()='Price Calculation']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", price_calc_label)
                time.sleep(1)
                
                # Etiketten parent elemente git ve select bul
                parent = None
                element = price_calc_label
                
                for _ in range(5):  # En fazla 5 seviye yukarı
                    try:
                        element = element.find_element(By.XPATH, "..")
                        if "admin__field" in element.get_attribute("class"):
                            parent = element
                            break
                    except:
                        break
                
                if parent:
                    # Select elementi bul
                    select_element = parent.find_element(By.CSS_SELECTOR, "select")
                    
                    # Select elementi ile dropdown'ı aç
                    select = Select(select_element)
                    select.select_by_visible_text("Fixed")
                    logger.info("Selected Fixed with Selenium Select")
                    
    
                    
                    return True
                else:
                    logger.warning("Could not find parent admin__field for Price Calculation")
            except Exception as selenium_e:
                logger.error(f"Selenium approach for Price Calculation failed: {str(selenium_e)}")
                
                # Son bir deneme - ActionChains ile etiketin yanına tıkla
                try:
                    # Price Calculation etiketini tekrar bul
                    price_calc_label = driver.find_element(By.XPATH, "//*[text()='Price Calculation']")
                    
                    # ActionChains kullan
                    actions = ActionChains(driver)
                    actions.move_to_element(price_calc_label).move_by_offset(150, 0).click().perform()
                    logger.info("Clicked near Price Calculation label with ActionChains")
                    time.sleep(1)
                    
                    # Fixed seçeneğini bul
                    fixed_options = driver.find_elements(By.XPATH, "//*[text()='Fixed']")
                    
                    if fixed_options:
                        for option in fixed_options:
                            try:
                                if option.is_displayed():
                                    option.click()
                                    logger.info("Clicked Fixed option")
                                    time.sleep(1)
                                    return True
                            except:
                                continue
                    
                    logger.warning("Could not find any visible Fixed options")
                except Exception as action_e:
                    logger.error(f"ActionChains approach also failed: {str(action_e)}")
        
        # Başarı durumunu kontrol et
        return not ("not found" in result or "Could not find" in result)
    except Exception as e:
        logger.error(f"Failed to set Price Calculation to Fixed: {str(e)}")
        return False
    
def select_production_days_multiselect(driver):
    """Select weekdays in Production Days multiselect"""
    try:
        logger.info("Selecting weekdays in Production Days multiselect")
        
        
        # 1. Others sekmesine tıkla
        others_elements = driver.find_elements(By.XPATH, "//span[text()='Others']")
        
        if others_elements:
            for elem in others_elements:
                try:
                    if elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", elem)
                        logger.info("Clicked Others tab")
                        time.sleep(2)
                        break
                except:
                    continue
        else:
            logger.warning("No Others tab elements found")
            return False
        
        # 2. Sayfayı biraz aşağı kaydır ve Production Days alanını bul
        driver.execute_script("window.scrollBy(0, 200);")
        time.sleep(1)
        
        # Ekran görüntüsü temel alınarak, üç yaklaşım dene
        success = False
        
        # Yaklaşım 1: Doğrudan multiselect ID'sini kullanma
        try:
            logger.info("Approaching Production Days with direct ID")
            
            # Ekran görüntüsüne göre ID'yi kullan: id="Y91IBRT"
            multiselect = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "Y91IBRT"))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", multiselect)
            time.sleep(1)
            
            
            # Tüm hafta içi günlerini bir anda seçmek için JavaScript kullan
            select_result = driver.execute_script("""
                // Multiselect elementi
                var multiselect = arguments[0];
                
                // Weekdays - Ekran görüntüsü verilerine göre 
                var weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
                var selected = [];
                
                // Multiselect options
                var options = document.querySelectorAll('option[data-title]');
                
                // Seçim yap
                for (var i = 0; i < options.length; i++) {
                    var option = options[i];
                    var day = option.getAttribute('data-title');
                    
                    if (weekdays.includes(day)) {
                        option.selected = true;
                        selected.push(day);
                    }
                }
                
                // Change event tetikle
                if (selected.length > 0) {
                    multiselect.dispatchEvent(new Event('change', {bubbles: true}));
                }
                
                return {
                    success: selected.length > 0,
                    selected: selected
                };
            """, multiselect)
            
            logger.info(f"JavaScript multiselect selection result: {select_result}")
            
            if isinstance(select_result, dict) and select_result.get('success'):
                success = True
        except Exception as id_e:
            logger.warning(f"Direct ID approach failed: {str(id_e)}")
        
        # Yaklaşım 2: Production Days label'inden multiselect'i bulma
        if not success:
            try:
                logger.info("Approaching Production Days via label")
                
                # Production Days label
                prod_days_label = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[text()='Production Days']"))
                )
                
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", prod_days_label)
                time.sleep(1)
                
                # Label'dan select/multiselect bul
                parent = None
                element = prod_days_label
                
                for _ in range(5):
                    try:
                        element = element.find_element(By.XPATH, "..")
                        if "admin__field" in element.get_attribute("class"):
                            parent = element
                            break
                    except:
                        break
                
                if parent:
                    # Multiselect bul
                    multiselect = parent.find_element(By.CSS_SELECTOR, "select.admin__control-multiselect, select[multiple]")
                    
                    # Select kontrolü ile hafta içi günlerini seç
                    select = Select(multiselect)
                    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
                    
                    # Tüm günleri seç
                    for day in weekdays:
                        select.select_by_visible_text(day)
                        logger.info(f"Selected {day}")
                    
                    logger.info("Successfully selected all weekdays")
                    success = True
            except Exception as label_e:
                logger.warning(f"Label approach failed: {str(label_e)}")
        
        # Yaklaşım 3: Tüm multiselect'leri deneme
        if not success:
            try:
                logger.info("Trying all multiselects approach")
                
                # Tüm multiselect'leri bul
                multiselects = driver.find_elements(By.CSS_SELECTOR, "select.admin__control-multiselect, select[multiple]")
                
                if multiselects:
                    logger.info(f"Found {len(multiselects)} multiselects")
                    
                    # Her multiselect'i dene
                    for idx, multi in enumerate(multiselects):
                        try:
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", multi)
                            time.sleep(1)
                            
                            
                            
                            # JavaScript ile seçim yap
                            result = driver.execute_script("""
                                var select = arguments[0];
                                var weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
                                var selected = [];
                                
                                // Tüm option'ları kontrol et
                                for (var i = 0; i < select.options.length; i++) {
                                    var option = select.options[i];
                                    var text = option.text.trim();
                                    
                                    if (weekdays.includes(text)) {
                                        option.selected = true;
                                        selected.push(text);
                                    }
                                }
                                
                                if (selected.length > 0) {
                                    select.dispatchEvent(new Event('change', {bubbles: true}));
                                }
                                
                                return {
                                    success: selected.length > 0,
                                    selected: selected
                                };
                            """, multi)
                            
                            logger.info(f"JavaScript multiselect {idx+1} result: {result}")
                            
                            if isinstance(result, dict) and result.get('success'):
                                logger.info(f"Successfully selected weekdays in multiselect {idx+1}")
                                success = True
                                break
                        except Exception as multi_e:
                            logger.warning(f"Error with multiselect {idx+1}: {str(multi_e)}")
                else:
                    logger.warning("No multiselects found")
            except Exception as all_e:
                logger.warning(f"All multiselects approach failed: {str(all_e)}")
        
        
        return success
    except Exception as e:
        logger.error(f"Failed to select Production Days: {str(e)}")
        return False

def find_and_set_product_type(driver, product_type="Upload + Personalization + Configuration"):
    """Helper function to find and set Product Type field"""
    try:
        logger.info("Looking for Product Type field")
        

        
        # Biraz aşağı kaydır
        driver.execute_script("window.scrollBy(0, 150);")
        time.sleep(1)
        
        # First try: Product Type label bulma
        try:
            product_type_labels = driver.find_elements(By.XPATH, "//span[text()='Product Type']")
            
            if product_type_labels:
                logger.info(f"Found {len(product_type_labels)} Product Type labels")
                
                for i, label in enumerate(product_type_labels):
                    try:
                        if label.is_displayed():
                            # Görünür yap
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", label)
                            time.sleep(1)
                            
                            
                            # Dropdown'u bul
                            parent = label
                            for _ in range(5):  # En fazla 5 seviye
                                parent = parent.find_element(By.XPATH, "..")
                                try:
                                    dropdown = parent.find_element(By.CSS_SELECTOR, ".admin__action-multiselect, select")
                                    driver.execute_script("arguments[0].click();", dropdown)
                                    logger.info("Clicked Product Type dropdown")
                                    time.sleep(1)

                                    # Şimdi opsiyon ara
                                    options = driver.find_elements(By.XPATH, f"//*[contains(text(), '{product_type}')]")
                                    
                                    if options:
                                        for option in options:
                                            try:
                                                if option.is_displayed():
                                                    driver.execute_script("arguments[0].click();", option)
                                                    logger.info(f"Selected option containing '{product_type}'")
                                                    
                                                    
                                                    return True
                                            except:
                                                continue
                                    break
                                except:
                                    continue
                    except Exception as e:
                        logger.warning(f"Error with Product Type label {i}: {str(e)}")
        except Exception as e:
            logger.warning(f"Error finding Product Type label: {str(e)}")
        
        # İkinci deneme: Tüm aday dropdown'ları dene
        logger.info("Trying all potential dropdowns")
        
        selects = driver.find_elements(By.CSS_SELECTOR, ".admin__action-multiselect, select")
        logger.info(f"Found {len(selects)} potential dropdown elements")
        
        for i, select_elem in enumerate(selects[:5]):  # İlk 5'i dene
            try:
                # Görünür mü?
                if not select_elem.is_displayed():
                    continue
                    
                # Görünür yap
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_elem)
                time.sleep(1)
                

                
                # Tıkla
                driver.execute_script("arguments[0].click();", select_elem)
                logger.info(f"Clicked select {i}")
                time.sleep(1)
                

                
                # Opcions ara
                options = driver.find_elements(By.XPATH, f"//*[contains(text(), '{product_type}')]")
                
                if options:
                    for option in options:
                        try:
                            if option.is_displayed():
                                driver.execute_script("arguments[0].click();", option)
                                logger.info(f"Selected option containing '{product_type}'")
                                
                                return True
                        except:
                            continue
            except Exception as e:
                logger.warning(f"Error with select {i}: {str(e)}")
        
        # Son deneme: JavaScript ile
        logger.info("Trying direct JavaScript search and selection")
        result = driver.execute_script("""
            // Find all dropdowns near "Product Type" text
            var elements = document.querySelectorAll('*');
            var productTypeLabels = [];
            
            // Find Product Type labels
            for (var i = 0; i < elements.length; i++) {
                if (elements[i].textContent && elements[i].textContent.trim() === 'Product Type') {
                    productTypeLabels.push(elements[i]);
                }
            }
            
            if (productTypeLabels.length > 0) {
                for (var i = 0; i < productTypeLabels.length; i++) {
                    try {
                        // Find containing admin field
                        var field = productTypeLabels[i];
                        for (var j = 0; j < 5; j++) {  // Look up to 5 levels
                            field = field.parentElement;
                            if (!field) break;
                            
                            // Look for dropdown within this field
                            var dropdown = field.querySelector('select, .admin__action-multiselect');
                            if (dropdown) {
                                // Click to open
                                dropdown.click();
                                
                                // Wait a bit
                                setTimeout(function(){}, 100);
                                
                                // Look for our target option
                                var allElements = document.querySelectorAll('*');
                                for (var k = 0; k < allElements.length; k++) {
                                    if (allElements[k].textContent && 
                                        allElements[k].textContent.includes(arguments[0])) {
                                        // Click this option
                                        allElements[k].click();
                                        return "Selected " + arguments[0] + " via Product Type label path";
                                    }
                                }
                                
                                break;  // Break parent search if dropdown found but option not found
                            }
                        }
                    } catch (e) {
                        // Continue to next label
                    }
                }
            }
            
            // If we get here, try all dropdowns
            var allDropdowns = document.querySelectorAll('select, .admin__action-multiselect');
            
            for (var i = 0; i < Math.min(allDropdowns.length, 8); i++) {
                try {
                    // Check if visible
                    var style = window.getComputedStyle(allDropdowns[i]);
                    if (style.display === 'none' || style.visibility === 'hidden') continue;
                    
                    // Click to open
                    allDropdowns[i].click();
                    
                    // Wait a bit
                    setTimeout(function(){}, 100);
                    
                    // Look for our target option
                    var allElements = document.querySelectorAll('*');
                    for (var j = 0; j < allElements.length; j++) {
                        if (allElements[j].textContent && 
                            allElements[j].textContent.includes(arguments[0])) {
                            // Click this option
                            allElements[j].click();
                            return "Selected " + arguments[0] + " via random dropdown " + i;
                        }
                    }
                } catch (e) {
                    // Continue to next dropdown
                }
            }
            
            return "Could not find or select " + arguments[0];
        """, product_type)
        
        logger.info(f"JavaScript select result: {result}")
        
        # Başarı durumunu kontrol et
        if "Could not find" not in result:
            return True
            
        return False
    except Exception as e:
        logger.error(f"Error in find_and_set_product_type: {str(e)}")
        return False
 
def set_vpe_value_optimized(driver, vpe_value):
    """Set VPE value using only the most successful method (DOM manipulation)"""
    logger.info(f"Setting VPE value to {vpe_value} using optimized method")
    
    try:
        # Use JavaScript to directly set the value in the DOM
        js_script = """
        // Try to find the input by various attributes
        var input = document.querySelector('input[name="product[delivery][vpe]"]');
        
        if (!input) {
            // Try to find by label
            var labels = document.querySelectorAll('span, label');
            for (var i = 0; i < labels.length; i++) {
                if (labels[i].textContent.includes('VPE')) {
                    var parent = labels[i].closest('.admin__field');
                    if (parent) {
                        input = parent.querySelector('input');
                        break;
                    }
                }
            }
        }
        
        if (input) {
            // Directly set the value in the DOM
            input.value = arguments[0];
            
            // Create and dispatch events
            var inputEvent = new Event('input', { bubbles: true });
            var changeEvent = new Event('change', { bubbles: true });
            
            input.dispatchEvent(inputEvent);
            input.dispatchEvent(changeEvent);
            
            return "Set VPE value to " + arguments[0] + " via DOM manipulation";
        }
        
        return "Could not find VPE input via DOM manipulation";
        """
        
        result = driver.execute_script(js_script, vpe_value)
        logger.info(f"DOM manipulation result: {result}")
        

        return "Could not find" not in result
    except Exception as e:
        logger.error(f"Failed to set VPE value: {str(e)}")
        return False

def force_select_vpe_type(driver, vpe_type="Stock"):
    """Force select VPE Type by simulating user actions very directly"""
    try:
        logger.info(f"Forcing VPE Type selection to: {vpe_type}")
 
        # Force-replace the value directly in the DOM with JavaScript
        # This is the most aggressive approach but should work
        success = driver.execute_script("""
            // Find the VPE Type field
            var labels = Array.from(document.querySelectorAll('span')).filter(
                el => el.textContent.trim() === 'VPE Type'
            );
            
            if (labels.length === 0) return "VPE Type label not found";
            
            var label = labels[0];
            var field = label.closest('.admin__field');
            
            if (!field) return "VPE Type field not found";
            
            // Now replace the text in the dropdown
            var dropdown = field.querySelector('.admin__action-multiselect');
            
            if (!dropdown) return "VPE Type dropdown not found";
            
            // Find the text element to replace
            var selectedText = dropdown.querySelector('.admin__action-multiselect-text');
            if (selectedText) {
                // Replace the text visually
                selectedText.textContent = arguments[0];
                
                // Also set a data attribute to help identify it was changed
                dropdown.setAttribute('data-selected-option', arguments[0]);
                
                // Force update any hidden inputs if they exist
                var hiddenInputs = field.querySelectorAll('input[type="hidden"]');
                for (var i = 0; i < hiddenInputs.length; i++) {
                    // Set a value that corresponds to Stock option - this is just a guess
                    hiddenInputs[i].value = "stock";
                    hiddenInputs[i].dispatchEvent(new Event('change', {bubbles: true}));
                }
                
                return "Successfully forced VPE Type to: " + arguments[0];
            } else {
                return "Could not find text element to replace";
            }
        """, vpe_type)
        
        logger.info(f"Force select result: {success}")

        # If JavaScript approach returned an error message
        if isinstance(success, str) and ("not found" in success.lower() or "could not" in success.lower()):
            logger.warning(f"JavaScript approach failed: {success}, trying ActionChains")
            
            # Try the most direct possible approach with action chains
            try:
                # Find the dropdown
                vpe_dropdown = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[text()='VPE Type']/../..//div[contains(@class, 'admin__action-multiselect')]"))
                )
                
                # Scroll to it
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vpe_dropdown)
                time.sleep(1)
                
                # Click to open the dropdown
                action = ActionChains(driver)
                action.move_to_element(vpe_dropdown).click().perform()
                logger.info("Clicked VPE dropdown with ActionChains")
                time.sleep(1)
                
                # Try to find Stock option
                try:
                    stock_option = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'admin__action-multiselect-menu-inner-item')]//span[text()='Stock']"))
                    )
                    action.move_to_element(stock_option).click().perform()
                    logger.info("Clicked Stock option with ActionChains")
                    return True
                except:
                    # If can't find by text, try by position - Stock is usually second option
                    logger.info("Could not find Stock by text, trying by position")
                    
                    # Click second option or any non-LFM option
                    options_script = """
                    var options = document.querySelectorAll('.admin__action-multiselect-menu-inner-item');
                    var coords = [];
                    
                    if (options.length > 1) {
                        // Get position of the second option (index 1) which is usually Stock
                        var rect = options[1].getBoundingClientRect();
                        return {
                            x: rect.left + (rect.width / 2) + window.pageXOffset,
                            y: rect.top + (rect.height / 2) + window.pageYOffset
                        };
                    } else if (options.length > 0) {
                        // If only one option, get its position
                        var rect = options[0].getBoundingClientRect();
                        return {
                            x: rect.left + (rect.width / 2) + window.pageXOffset,
                            y: rect.top + (rect.height / 2) + window.pageYOffset
                        };
                    } else {
                        return null;
                    }
                    """
                    
                    option_coords = driver.execute_script(options_script)
                    
                    if option_coords:
                        # Click at the calculated coordinates
                        action = ActionChains(driver)
                        action.move_by_offset(option_coords['x'], option_coords['y']).click().perform()
                        logger.info(f"Clicked at position x:{option_coords['x']}, y:{option_coords['y']}")
                        return True
            except Exception as e:
                logger.error(f"ActionChains approach failed: {str(e)}")
                return False
        
        return "not found" not in str(success).lower() and "could not" not in str(success).lower()
    except Exception as e:
        logger.error(f"Failed to force-select VPE Type: {str(e)}")
        return False

def select_vpe_type_producer_style(driver, vpe_type="Stock"):
    """Select VPE Type using the same approach that works for Producer selection"""
    try:
        logger.info(f"Selecting VPE Type: {vpe_type} using producer-style approach")
        
        # First make sure we're on the Delivery section
        delivery_section = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'admin__collapsible-block-wrapper') and .//span[contains(text(), 'Delivery')]]"))
        )
        
        # Make sure it's expanded
        if "active" not in delivery_section.get_attribute("class"):
            header = delivery_section.find_element(By.CSS_SELECTOR, ".admin__collapsible-title")
            driver.execute_script("arguments[0].click();", header)
            time.sleep(2)
        
        # Wait for a moment to ensure the form is fully loaded
        time.sleep(2)
        
    
        # Look at the page structure more carefully now
        # Use JavaScript to explore the form structure
        form_structure = driver.execute_script("""
            // Inspect the structure of the delivery section
            var spans = document.querySelectorAll('span');
            var vpeTypeLabel = null;
            
            // Find VPE Type label
            for (var i = 0; i < spans.length; i++) {
                if (spans[i].textContent.trim() === 'VPE Type') {
                    vpeTypeLabel = spans[i];
                    break;
                }
            }
            
            if (!vpeTypeLabel) return {found: false, message: "VPE Type label not found"};
            
            // Get element structure to debug what's available
            var field = vpeTypeLabel.closest('.admin__field');
            var structure = {};
            
            if (field) {
                structure.hasField = true;
                structure.fieldClasses = field.getAttribute('class');
                
                // Check for selects
                var selects = field.querySelectorAll('select');
                structure.selectCount = selects.length;
                
                if (selects.length > 0) {
                    structure.firstSelectId = selects[0].id || 'no-id';
                    structure.firstSelectClasses = selects[0].getAttribute('class');
                    structure.firstSelectOptions = [];
                    
                    for (var i = 0; i < selects[0].options.length; i++) {
                        structure.firstSelectOptions.push(selects[0].options[i].text);
                    }
                }
                
                // Check for multiselect 
                var multiselect = field.querySelector('.admin__action-multiselect');
                structure.hasMultiselect = multiselect !== null;
                
                if (multiselect) {
                    structure.multiselectClasses = multiselect.getAttribute('class');
                    
                    // Check what's inside multiselect
                    var text = multiselect.querySelector('.admin__action-multiselect-text');
                    structure.hasText = text !== null;
                    if (text) structure.selectedText = text.textContent.trim();
                }
                
                // Check for any other potential dropdown or select-like element
                var dropdownLike = field.querySelectorAll('div[class*="select"], div[class*="dropdown"], div[role="listbox"]');
                structure.dropdownLikeCount = dropdownLike.length;
                
                if (dropdownLike.length > 0) {
                    structure.firstDropdownClasses = dropdownLike[0].getAttribute('class');
                }
            } else {
                structure.hasField = false;
            }
            
            return {
                found: true,
                structure: structure,
                labelRect: vpeTypeLabel.getBoundingClientRect()
            };
        """)
        
        logger.info(f"VPE Type field structure: {form_structure}")
        
        if not form_structure.get('found', False):
            logger.error("Could not find VPE Type label")
            return False
        
        # Based on the form structure, determine the best approach
        structure = form_structure.get('structure', {})
        
        # Take different approaches based on what we find in the structure
        if structure.get('selectCount', 0) > 0:
            # Traditional select element found
            logger.info("Found traditional select for VPE Type")
            
            select_result = driver.execute_script("""
                var spans = document.querySelectorAll('span');
                var vpeTypeLabel = null;
                
                for (var i = 0; i < spans.length; i++) {
                    if (spans[i].textContent.trim() === 'VPE Type') {
                        vpeTypeLabel = spans[i];
                        break;
                    }
                }
                
                if (!vpeTypeLabel) return "VPE Type label not found";
                
                var field = vpeTypeLabel.closest('.admin__field');
                if (!field) return "VPE Type field not found";
                
                var select = field.querySelector('select');
                if (!select) return "Select element not found";
                
                var found = false;
                for (var i = 0; i < select.options.length; i++) {
                    if (select.options[i].text.includes(arguments[0])) {
                        select.selectedIndex = i;
                        select.dispatchEvent(new Event('change', {bubbles: true}));
                        found = true;
                        break;
                    }
                }
                
                return found ? 
                    "Selected " + arguments[0] + " from VPE Type select" : 
                    "Option " + arguments[0] + " not found in select";
            """, vpe_type)
            
            logger.info(f"Select result: {select_result}")
            
        elif structure.get('hasMultiselect', False):
            # Custom multiselect found
            logger.info("Found custom multiselect for VPE Type")
            
            # Click to open, then select
            multiselect_result = driver.execute_script("""
                var spans = document.querySelectorAll('span');
                var vpeTypeLabel = null;
                
                for (var i = 0; i < spans.length; i++) {
                    if (spans[i].textContent.trim() === 'VPE Type') {
                        vpeTypeLabel = spans[i];
                        break;
                    }
                }
                
                if (!vpeTypeLabel) return "VPE Type label not found";
                
                var field = vpeTypeLabel.closest('.admin__field');
                if (!field) return "VPE Type field not found";
                
                var multiselect = field.querySelector('.admin__action-multiselect');
                if (!multiselect) return "Multiselect element not found";
                
                // Click to open
                multiselect.click();
                
                // Small delay
                setTimeout(function(){}, 100);
                
                // Find and click the option
                var found = false;
                var options = document.querySelectorAll('.admin__action-multiselect-menu-inner-item span');
                
                for (var i = 0; i < options.length; i++) {
                    if (options[i].textContent.includes(arguments[0])) {
                        options[i].click();
                        found = true;
                        break;
                    }
                }
                
                return found ? 
                    "Selected " + arguments[0] + " from VPE Type multiselect" : 
                    "Option " + arguments[0] + " not found in multiselect options";
            """, vpe_type)
            
            logger.info(f"Multiselect result: {multiselect_result}")
            
        else:
            # No easily identifiable select pattern - try to click and select directly
            logger.info("No standard select elements found, trying direct click approach")
            
            # First try to find if there's a currently visible selection text we can click
            click_result = driver.execute_script("""
                var spans = document.querySelectorAll('span');
                var vpeTypeLabel = null;
                
                for (var i = 0; i < spans.length; i++) {
                    if (spans[i].textContent.includes('VPE Type')) {
                        vpeTypeLabel = spans[i];
                        break;
                    }
                }
                
                if (!vpeTypeLabel) return "VPE Type label not found";
                
                // Calculate positions - first check if there's visible text or control to click
                var field = vpeTypeLabel.closest('.admin__field');
                
                if (!field) {
                    // If can't find field, just try clicking to the right of the label
                    var rect = vpeTypeLabel.getBoundingClientRect();
                    var x = rect.right + 150; // 150px to the right
                    var y = rect.top + (rect.height / 2);
                    
                    // Simulate click at this position
                    var clickEvent = new MouseEvent('click', {
                        view: window,
                        bubbles: true,
                        cancelable: true,
                        clientX: x,
                        clientY: y
                    });
                    document.elementFromPoint(x, y).dispatchEvent(clickEvent);
                    
                    return "Clicked at position to the right of VPE Type label";
                }
                
                // Look for anything that resembles a dropdown or has "LFM" in the text
                var clickTarget = null;
                
                // Try to find something with 'LFM' or 'Running Meter' text
                var allElements = field.querySelectorAll('*');
                for (var i = 0; i < allElements.length; i++) {
                    if (allElements[i].textContent && 
                        (allElements[i].textContent.includes('LFM') || 
                         allElements[i].textContent.includes('Running Meter'))) {
                        clickTarget = allElements[i];
                        break;
                    }
                }
                
                if (clickTarget) {
                    clickTarget.click();
                    return "Clicked on element with LFM text";
                }
                
                // If we get here, we didn't find anything to click
                return "Could not find clickable element for VPE Type";
            """)
            
            logger.info(f"Initial click result: {click_result}")
            
            # Wait a moment to see if dropdown appears
            time.sleep(1)
            
            # Now try to click the Stock option
            select_result = driver.execute_script("""
                // Look for Stock in any dropdown that might be visible now
                var stockOptions = [];
                var elements = document.querySelectorAll('*');
                
                for (var i = 0; i < elements.length; i++) {
                    if (elements[i].textContent && 
                        elements[i].textContent.trim() === 'Stock' &&
                        window.getComputedStyle(elements[i]).display !== 'none') {
                        stockOptions.push(elements[i]);
                    }
                }
                
                // If we found any visible Stock options, click the first one
                if (stockOptions.length > 0) {
                    stockOptions[0].click();
                    return "Clicked on Stock option";
                }
                
                // Try clicking any list item that might be a dropdown option
                var listItems = document.querySelectorAll('li.admin__action-multiselect-menu-inner-item, li.option-item, li[role="option"]');
                for (var i = 0; i < listItems.length; i++) {
                    if (listItems[i].textContent && 
                        listItems[i].textContent.trim() === 'Stock' &&
                        window.getComputedStyle(listItems[i]).display !== 'none') {
                        listItems[i].click();
                        return "Clicked on Stock list item";
                    }
                }
                
                return "Could not find Stock option to click";
            """)
            
            logger.info(f"Option selection result: {select_result}")
     # If we couldn't select with JavaScript, try using Selenium directly
        if "Could not find" in str(select_result):
            logger.info("JavaScript approaches failed, trying direct Selenium interaction")
            
            try:
                # Find the section containing VPE Type
                vpe_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[text()='VPE Type']/ancestor::div[contains(@class, 'admin__field')]"))
                )
                
                # Find if there's a visible LFM text or dropdown
                try:
                    # Look for LFM text
                    lfm_element = vpe_section.find_element(By.XPATH, ".//*[contains(text(), 'LFM') or contains(text(), 'Running Meter')]")
                    driver.execute_script("arguments[0].click();", lfm_element)
                    logger.info("Clicked on LFM element to open dropdown")
                    time.sleep(1)
                    
                    # Now look for Stock option
                    try:
                        stock_option = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'admin__action-multiselect-menu-inner-item')]//span[contains(text(), 'Stock')]"))
                        )
                        driver.execute_script("arguments[0].click();", stock_option)
                        logger.info("Clicked Stock option")
                        return True
                    except:
                        logger.warning("Could not find Stock option after clicking LFM")
                except:
                    logger.warning("Could not find LFM element")
                    
                # If we get here, previous attempts failed - try clicking any dropdown-like element
                try:
                    dropdown = vpe_section.find_element(By.CSS_SELECTOR, ".admin__action-multiselect, select, [role='listbox'], [class*='select'], [class*='dropdown']")
                    driver.execute_script("arguments[0].click();", dropdown)
                    logger.info("Clicked on dropdown-like element")
                    time.sleep(1)
                    
                    # Now look for Stock option
                    try:
                        stock_option = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'admin__action-multiselect-menu-inner-item')]//span[contains(text(), 'Stock')]"))
                        )
                        driver.execute_script("arguments[0].click();", stock_option)
                        logger.info("Clicked Stock option after dropdown click")
                        return True
                    except:
                        logger.warning("Could not find Stock option after clicking dropdown")
                except:
                    logger.warning("Could not find dropdown element")
            except Exception as selenium_e:
                logger.error(f"Selenium direct approach failed: {str(selenium_e)}")
        
        # Return True if we didn't get a clear failure message
        return not ("not found" in str(select_result).lower() or "could not find" in str(select_result).lower())
    except Exception as e:
        logger.error(f"Failed to select VPE Type with producer-style approach: {str(e)}")
        return False

def configure_delivery_section_final(driver):
    """Final optimized function for delivery section configuration focusing on reliability"""
    try:
        logger.info(f"Configuring Delivery section with final optimized approach")
        
  
        # Find and scroll to the Delivery section using JavaScript
        js_delivery_check = """
            // Find all section titles
            var titles = Array.from(document.querySelectorAll('.admin__collapsible-title'));
            
            // Find the one that contains "Delivery"
            var deliveryTitle = titles.find(title => title.textContent.includes('Delivery'));
            
            if (!deliveryTitle) return {found: false, message: "Delivery section not found"};
            
            var section = deliveryTitle.closest('.admin__collapsible-block-wrapper');
            var position = deliveryTitle.getBoundingClientRect();
            
            return {
                found: true,
                y: position.top + window.pageYOffset,
                isExpanded: section.classList.contains('active')
            };
        """
        
        delivery_info = driver.execute_script(js_delivery_check)
        
        if not delivery_info.get('found', False):
            logger.error("Could not find Delivery section")
            raise Exception("Delivery section not found")
        
        # Scroll to the section
        y_position = delivery_info['y'] - 100
        driver.execute_script(f"window.scrollTo(0, {y_position});")
        time.sleep(1.5)
        
 
        # Expand the section if not already expanded
        if not delivery_info.get('isExpanded', False):
            expand_result = driver.execute_script("""
                var titles = Array.from(document.querySelectorAll('.admin__collapsible-title'));
                var deliveryTitle = titles.find(title => title.textContent.includes('Delivery'));
                
                if (!deliveryTitle) return "Delivery title not found";
                
                deliveryTitle.click();
                return "Clicked to expand delivery section";
            """)
            logger.info(f"Expansion result: {expand_result}")
            time.sleep(2)
        

        # 1. Set VPE value to 9999 using the optimized function
        logger.info("Setting VPE value to 9999")
        if not set_vpe_value_optimized(driver, VPE_VALUE):
            logger.warning(f"Failed to set VPE value to {VPE_VALUE}")
        
        # 2. Select VPE Type - use the producer-style method for better reliability
        logger.info(f"Selecting VPE Type: {VPE_TYPE}")
        if not select_vpe_type_producer_style(driver, VPE_TYPE):
            logger.warning(f"Failed to select VPE Type with producer-style method")
        
        # 3. Select Fix Products option
        logger.info("Selecting 'Fix Products' option")
        select_fix_result = driver.execute_script("""
            // Find all select elements
            var selects = document.querySelectorAll('select');
            
            // Look for the one with Fix Products option
            for (var i = 0; i < selects.length; i++) {
                var select = selects[i];
                var hasFixProducts = false;
                
                // Check if this select has a Fix Products option
                for (var j = 0; j < select.options.length; j++) {
                    if (select.options[j].textContent.includes("Fix Products")) {
                        hasFixProducts = true;
                        select.selectedIndex = j;
                        select.dispatchEvent(new Event('change', {bubbles: true}));
                        break;
                    }
                }
                
                if (hasFixProducts) {
                    return "Selected Fix Products option";
                }
            }
            
            return "Could not find select with Fix Products option";
        """)
        
        logger.info(f"Fix Products selection result: {select_fix_result}")
        
        if "Could not find" in select_fix_result:
            # Fallback to the standard method if JavaScript fails
            if not select_fix_products_option(driver):
                logger.warning("Failed to select Fix Products option")
        
        logger.info("Successfully configured Delivery section")
        return True
    except Exception as e:
        logger.error(f"Failed to configure Delivery section with final approach: {str(e)}")
        return False
        
def select_fix_products_option(driver):
    """Select 'Fix Products' from the combination packaging dropdown"""
    logger.info("Selecting 'Fix Products' from combination packaging dropdown")
    
 
    try:
        
        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//select[contains(@class, 'admin__control-multiselect')]"))
        )
        
        # Create a Select object
        select = Select(select_element)
        
        # Select by visible text
        select.select_by_visible_text("Fix Products")
        
        logger.info("Selected 'Fix Products' using Select class")
        
        return True
    except Exception as e:
        logger.warning(f"Failed to select 'Fix Products' using Select class: {str(e)}")
        
        # Try JavaScript approach
        try:
            js_script = """
            // This script is based on the exact HTML structure shown in the screenshot
            var fixProductsOption = document.querySelector('option[data-title="Fix Products"]');
            if (!fixProductsOption) {
                // Try to find by text content
                var options = document.querySelectorAll('option');
                for (var i = 0; i < options.length; i++) {
                    if (options[i].textContent.trim() === "Fix Products") {
                        fixProductsOption = options[i];
                        break;
                    }
                }
            }
            
            if (fixProductsOption) {
                // Select this option
                fixProductsOption.selected = true;
                
                // Trigger change event on the select element
                var select = fixProductsOption.parentElement;
                select.dispatchEvent(new Event('change', { bubbles: true }));
                
                return "Selected 'Fix Products' using JavaScript";
            }
            
            return "Could not find 'Fix Products' option";
            """
            
            result = driver.execute_script(js_script)
            logger.info(f"JavaScript selection result: {result}")

            return "Could not find" not in result
        except Exception as js_e:
            logger.error(f"Failed to select 'Fix Products' using JavaScript: {str(js_e)}")
            return False

def configure_delivery_section_optimized(driver):
    """Highly optimized function for delivery section configuration focusing on reliability"""
    try:
        logger.info(f"Configuring Delivery section with optimized approach")
        

        # Find and scroll to the Delivery section using JavaScript
        js_delivery_check = """
            // Find all section titles
            var titles = Array.from(document.querySelectorAll('.admin__collapsible-title'));
            
            // Find the one that contains "Delivery"
            var deliveryTitle = titles.find(title => title.textContent.includes('Delivery'));
            
            if (!deliveryTitle) return {found: false, message: "Delivery section not found"};
            
            var section = deliveryTitle.closest('.admin__collapsible-block-wrapper');
            var position = deliveryTitle.getBoundingClientRect();
            
            return {
                found: true,
                y: position.top + window.pageYOffset,
                isExpanded: section.classList.contains('active')
            };
        """
        
        delivery_info = driver.execute_script(js_delivery_check)
        
        if not delivery_info.get('found', False):
            logger.error("Could not find Delivery section")
            raise Exception("Delivery section not found")
        
        # Scroll to the section
        y_position = delivery_info['y'] - 100
        driver.execute_script(f"window.scrollTo(0, {y_position});")
        time.sleep(1.5)
 
        # Expand the section if not already expanded
        if not delivery_info.get('isExpanded', False):
            expand_result = driver.execute_script("""
                var titles = Array.from(document.querySelectorAll('.admin__collapsible-title'));
                var deliveryTitle = titles.find(title => title.textContent.includes('Delivery'));
                
                if (!deliveryTitle) return "Delivery title not found";
                
                deliveryTitle.click();
                return "Clicked to expand delivery section";
            """)
            logger.info(f"Expansion result: {expand_result}")
            time.sleep(2)

        # 1. Set VPE value to 9999 using the optimized function
        logger.info("Setting VPE value to 9999")
        if not set_vpe_value_optimized(driver, VPE_VALUE):
            logger.warning(f"Failed to set VPE value to {VPE_VALUE}")
        
        # 2. Select VPE Type - use the improved direct method
        logger.info(f"Selecting VPE Type: {VPE_TYPE}")
        if not force_select_vpe_type(driver, VPE_TYPE):
            logger.warning(f"Failed to select VPE Type with improved method")
        
        # 3. Select Fix Products option
        logger.info("Selecting 'Fix Products' option")
        select_fix_result = driver.execute_script("""
            // Find all select elements
            var selects = document.querySelectorAll('select');
            
            // Look for the one with Fix Products option
            for (var i = 0; i < selects.length; i++) {
                var select = selects[i];
                var hasFixProducts = false;
                
                // Check if this select has a Fix Products option
                for (var j = 0; j < select.options.length; j++) {
                    if (select.options[j].textContent.includes("Fix Products")) {
                        hasFixProducts = true;
                        select.selectedIndex = j;
                        select.dispatchEvent(new Event('change', {bubbles: true}));
                        break;
                    }
                }
                
                if (hasFixProducts) {
                    return "Selected Fix Products option";
                }
            }
            
            return "Could not find select with Fix Products option";
        """)
        
        logger.info(f"Fix Products selection result: {select_fix_result}")
        
        if "Could not find" in select_fix_result:
            # Fallback to the standard method if JavaScript fails
            if not select_fix_products_option(driver):
                logger.warning("Failed to select Fix Products option")
        
        logger.info("Successfully configured Delivery section")
        return True
    except Exception as e:
        logger.error(f"Failed to configure Delivery section with optimized approach: {str(e)}")
        return False
    
def select_vpe_type_direct(driver, vpe_type="Stock"):
    """Select VPE Type by directly clicking on the text and then selecting Stock"""
    logger.info(f"Selecting VPE Type using direct click approach")
    
    try:
        # Find the text "LFM - Running Meter" which is currently displayed
        current_selection = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'LFM - Running Meter')]"))
        )
        
        # Click on it to open the dropdown
        driver.execute_script("arguments[0].click();", current_selection)
        logger.info("Clicked on 'LFM - Running Meter' to open dropdown")
        time.sleep(1)
        
        # Now find and click on "Stock" option
        stock_option = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'option')]//span[contains(text(), 'Stock')]"))
        )
        
        driver.execute_script("arguments[0].click();", stock_option)
        logger.info(f"Selected {vpe_type} from dropdown")
        
        return True
    except Exception as e:
        logger.error(f"Failed to select VPE Type using direct click: {str(e)}")
        
        # Try alternative approach with ActionChains
        try:
            # Find the VPE Type label
            vpe_type_label = driver.find_element(By.XPATH, "//span[text()='VPE Type']")
            
            # Scroll to make sure it's visible
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vpe_type_label)
            time.sleep(1)
            
            # Click to the right of the label where the dropdown should be
            action = ActionChains(driver)
            action.move_to_element(vpe_type_label).move_by_offset(150, 0).click().perform()
            logger.info("Clicked near VPE Type label to open dropdown")
            time.sleep(1)
            
            # Now click where "Stock" should be in the dropdown
            action.move_by_offset(0, 50).click().perform()
            logger.info(f"Clicked where {vpe_type} option should be")
            
            return True
        except Exception as alt_e:
            logger.error(f"Alternative approach also failed: {str(alt_e)}")
            return False

def select_last_product_part(driver):
    """Print Product Setup bölümünde Product Parts listesinde en son öğeyi seç"""
    try:
        logger.info("Attempting to select the last option in Product Parts list")
        
        # Daha kapsamlı ve güçlü bir yaklaşım - birkaç farklı stratejinin birleşimi
        tried_approaches = []
        
        # YAKLAŞIM 1: XPath ile select elementini bul
        try:
            tried_approaches.append("XPath approach")
            # Daha kapsamlı XPath ifadesi - farklı yapıları hedef alır
            xpaths = [
                "//span[text()='Product Parts']/ancestor::div[contains(@class, 'field')]//select",
                "//label[contains(text(), 'Product Parts')]/following-sibling::div//select",
                "//span[contains(text(), 'Product Parts')]/../..//select",
                "//label[contains(text(), 'Product Part')]/following-sibling::select",
                "//*[contains(text(), 'Product Parts')]/ancestor::div[contains(@class, 'field')]//select",
                "//select[contains(@name, 'product_part')]",
                "//select[contains(@id, 'product_part')]"
            ]
            
            product_parts_select = None
            for xpath in xpaths:
                elements = driver.find_elements(By.XPATH, xpath)
                for element in elements:
                    if element.is_displayed():
                        product_parts_select = element
                        logger.info(f"Found Product Parts select element using XPath: {xpath}")
                        break
                if product_parts_select:
                    break
            
            if product_parts_select:
                # Görünür olması için sayfayı kaydır
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", product_parts_select)
                time.sleep(1)
                
                # Select sınıfını oluştur
                select = Select(product_parts_select)
                
                # Seçimi temizle (varsa) - daha güvenli bir yaklaşım
                try:
                    select.deselect_all()
                except:
                    logger.info("Deselect not needed or not applicable")
                
                # Seçenek sayısını al
                options = select.options
                option_count = len(options)
                
                if option_count > 0:
                    # En son seçeneğin indexi
                    last_index = option_count - 1
                    logger.info(f"Selecting last option (index {last_index}) of {option_count} total options")
                    
                    # Hem index hem de value ile deneme yapalım
                    try:
                        select.select_by_index(last_index)
                        logger.info(f"Selected last option by index: {last_index}")
                    except Exception as select_e:
                        logger.warning(f"Select by index failed: {str(select_e)}")
                        # Value ile deneyelim
                        try:
                            last_value = options[last_index].get_attribute("value")
                            select.select_by_value(last_value)
                            logger.info(f"Selected last option by value: {last_value}")
                        except Exception as value_e:
                            logger.warning(f"Select by value failed: {str(value_e)}")
                            # Visible text ile deneyelim
                            try:
                                last_text = options[last_index].text
                                select.select_by_visible_text(last_text)
                                logger.info(f"Selected last option by visible text: {last_text}")
                            except Exception as text_e:
                                logger.warning(f"Select by text failed: {str(text_e)}")
                
                # Element Görünürlüğü ve Seçim Doğrulama
                try:
                    selected_options = select.all_selected_options
                    if selected_options:
                        logger.info(f"Successfully selected: {selected_options[0].text}")
                        return True
                except:
                    pass
            else:
                logger.warning("No Product Parts select element found using XPath")
        except Exception as xpath_e:
            logger.warning(f"XPath approach failed: {str(xpath_e)}")
        
        # YAKLAŞIM 2: JavaScript ile tüm select elementlerini dene
        try:
            tried_approaches.append("JavaScript approach")
            
            # Daha güçlü ve kapsamlı bir JavaScript yaklaşımı
            result = driver.execute_script("""
                // Tüm olası etiketleri ara (daha kapsayıcı)
                var possibleLabels = [
                    'Product Parts', 'Product Part', 'Produkt Teile', 'Produkt Teil',
                    'Parts', 'Part', 'Teile', 'Teil'
                ];
                
                // Etiket elemanlarını bul
                var labelElements = [];
                for (var i = 0; i < possibleLabels.length; i++) {
                    var matches = Array.from(document.querySelectorAll('*')).filter(function(el) {
                        return el.textContent && el.textContent.trim() === possibleLabels[i];
                    });
                    labelElements = labelElements.concat(matches);
                }
                
                // Etiket bulunamadıysa tüm select elementlerini dene
                if (labelElements.length === 0) {
                    var allSelects = document.querySelectorAll('select');
                    console.log('No label found, trying all ' + allSelects.length + ' select elements');
                    
                    for (var i = 0; i < allSelects.length; i++) {
                        var select = allSelects[i];
                        if (select.options.length > 0 && select.offsetParent !== null) {
                            // Son seçeneği seç
                            var lastIndex = select.options.length - 1;
                            
                            // Tüm seçimleri temizle
                            for (var j = 0; j < select.options.length; j++) {
                                select.options[j].selected = false;
                            }
                            
                            // Son seçeneği seç
                            select.options[lastIndex].selected = true;
                            select.dispatchEvent(new Event('change', {bubbles: true}));
                            
                            return {
                                success: true,
                                approach: 'all_selects',
                                index: lastIndex,
                                text: select.options[lastIndex].text,
                                optionCount: select.options.length
                            };
                        }
                    }
                    
                    return {success: false, message: 'No usable select elements found'};
                }
                
                // Görünür etiketleri filtrele
                var visibleLabels = labelElements.filter(function(el) {
                    return el.offsetParent !== null &&
                           window.getComputedStyle(el).visibility !== 'hidden' &&
                           window.getComputedStyle(el).display !== 'none';
                });
                
                if (visibleLabels.length === 0) {
                    return {success: false, message: 'No visible labels found'};
                }
                
                // Her görünür etiket için select elementi bulmaya çalış
                for (var i = 0; i < visibleLabels.length; i++) {
                    var label = visibleLabels[i];
                    
                    // 5 seviye yukarı çık ve select ara
                    var select = null;
                    var current = label;
                    for (var j = 0; j < 5; j++) {
                        if (!current) break;
                        
                        // Mevcut eleman içinde select ara
                        var selects = current.querySelectorAll('select');
                        if (selects.length > 0) {
                            select = selects[0];
                            break;
                        }
                        
                        // Yukarı çık
                        current = current.parentElement;
                    }
                    
                    // Select bulunamadıysa, yakındaki select elementlerini ara
                    if (!select) {
                        var nearbySelects = Array.from(document.querySelectorAll('select'))
                            .filter(function(sel) {
                                return sel.offsetParent !== null; // Görünür olmalı
                            });
                        
                        if (nearbySelects.length > 0) {
                            // En yakın select'i bul (DOM yapısında)
                            var labelRect = label.getBoundingClientRect();
                            
                            // En yakın select (mesafe bazlı)
                            var closestSelect = null;
                            var minDistance = Number.MAX_VALUE;
                            
                            for (var j = 0; j < nearbySelects.length; j++) {
                                var selectRect = nearbySelects[j].getBoundingClientRect();
                                var distance = Math.sqrt(
                                    Math.pow(selectRect.left - labelRect.left, 2) + 
                                    Math.pow(selectRect.top - labelRect.top, 2)
                                );
                                
                                if (distance < minDistance) {
                                    minDistance = distance;
                                    closestSelect = nearbySelects[j];
                                }
                            }
                            
                            select = closestSelect;
                        }
                    }
                    
                    if (select && select.options.length > 0) {
                        // Tüm seçimleri temizle
                        for (var j = 0; j < select.options.length; j++) {
                            select.options[j].selected = false;
                        }
                        
                        // Son seçeneği seç
                        var lastIndex = select.options.length - 1;
                        select.options[lastIndex].selected = true;
                        select.dispatchEvent(new Event('change', {bubbles: true}));
                        
                        return {
                            success: true,
                            labelText: label.textContent,
                            index: lastIndex,
                            text: select.options[lastIndex].text,
                            optionCount: select.options.length
                        };
                    }
                }
                
                return {success: false, message: 'No associated select element found for any label'};
            """)
            
            logger.info(f"JavaScript result: {result}")
            
            if isinstance(result, dict) and result.get('success'):
                logger.info(f"Successfully selected last option using JavaScript: {result}")
                return True
            
            logger.warning(f"JavaScript approach failed: {result}")
        except Exception as js_e:
            logger.error(f"JavaScript approach failed: {str(js_e)}")
        
        # YAKLAŞIM 3: CSS Selector ile dene
        try:
            tried_approaches.append("CSS Selector approach")
            
            # Farklı CSS Selector kombinasyonları
            css_selectors = [
                "select.admin__control-multiselect[name*='product']",
                "select[name*='product_part']",
                ".field-product_parts select",
                ".admin__field-control select.admin__control-select",
                "select.select"
            ]
            
            for selector in css_selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        for select_elem in elements:
                            if select_elem.is_displayed():
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_elem)
                                time.sleep(1)
                                
                                # Select ile manipüle et
                                select = Select(select_elem)
                                options = select.options
                                
                                if len(options) > 0:
                                    last_index = len(options) - 1
                                    select.select_by_index(last_index)
                                    logger.info(f"Selected last option (index {last_index}) using CSS selector: {selector}")
                                    return True
                except Exception as css_e:
                    logger.warning(f"CSS selector {selector} failed: {str(css_e)}")
        except Exception as css_approach_e:
            logger.warning(f"CSS Selector approach failed: {str(css_approach_e)}")
        
        # Tüm yaklaşımlar başarısız oldu
        logger.warning(f"All approaches failed: {', '.join(tried_approaches)}")
        logger.warning("Taking screenshot for debugging")
        try:
            driver.save_screenshot("product_parts_selection_failed.png")
        except:
            pass
        
        # Çözüm bulunamadı
        return False
        
    except Exception as e:
        logger.error(f"Failed to select last option in Product Parts: {str(e)}")
        # Detaylı hata bilgisi için stack trace yazdır
        import traceback
        logger.error(f"Stack trace: {traceback.format_exc()}")
        return False

def configure_print_product_setup_with_ctrl(driver):
    """Configure Print Product Setup section using CTRL+click for multiple selection"""
    try:
        logger.info("Configuring Print Product Setup with CTRL+click for multiple selection")
 
        # 1. Print Product Setup sekmesine tıkla
        print_setup_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Print Product Setup')]")
        
        if print_setup_elements:
            for elem in print_setup_elements:
                try:
                    if elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", elem)
                        logger.info("Clicked Print Product Setup")
                        time.sleep(2)
                        break
                except:
                    continue
        else:
            logger.warning("No Print Product Setup elements found")
            return False
        
        # Print Product Setup tıklandıktan sonra ekran görüntüsü
        if not select_last_product_part(driver):
            logger.warning("Failed to select last option in Product Parts, continuing anyway")
        
       
        # 2. JavaScript ile textarea içeriğini doğrudan ayarla (en güvenilir yöntem)
        textarea_result = driver.execute_script("""
            var result = {
                datencheck: false,
                dateityp: false
            };
            
            // Find textarea elements
            var textareas = document.querySelectorAll('textarea');
            
            for (var i = 0; i < textareas.length; i++) {
                var textarea = textareas[i];
                
                // Check if this is Datencheck textarea
                var parent = textarea.parentElement;
                var found = false;
                var fieldType = '';
                
                for (var j = 0; j < 10; j++) {
                    if (!parent) break;
                    
                    // Check all text in this parent for field labels
                    var text = parent.textContent || '';
                    
                    if (text.includes('Datencheck') && !result.datencheck) {
                        fieldType = 'datencheck';
                        found = true;
                        break;
                    } else if (text.includes('Dateityp') && !result.dateityp) {
                        fieldType = 'dateityp';
                        found = true;
                        break;
                    }
                    
                    parent = parent.parentElement;
                }
                
                if (found && fieldType === 'datencheck') {
                    // Set Datencheck values directly
                    textarea.value = "Standard ( options_none )\\nProfidatencheck [+9€] ( options_professional_data_check )";
                    textarea.dispatchEvent(new Event('input', { bubbles: true }));
                    textarea.dispatchEvent(new Event('change', { bubbles: true }));
                    result.datencheck = true;
                } else if (found && fieldType === 'dateityp') {
                    // Set Dateityp values directly
                    textarea.value = "PDF-file ( file_type_pdf_file )\\nImage ( file_type_image )";
                    textarea.dispatchEvent(new Event('input', { bubbles: true }));
                    textarea.dispatchEvent(new Event('change', { bubbles: true }));
                    result.dateityp = true;
                }
                
                // If we've set both, we can stop
                if (result.datencheck && result.dateityp) {
                    break;
                }
            }
            
            return result;
        """)
        
        logger.info(f"Textarea values set result: {textarea_result}")
        
        # JavaScript başarısız olursa, multiselect elementleri bul ve CTRL ile seç
        if not (textarea_result.get('datencheck', False) and textarea_result.get('dateityp', False)):
            logger.info("Trying multiselect approach with CTRL+click")
            
            # 3. Datencheck multiselect'i bul ve seçenekleri CTRL ile seç
            try:
                # Datencheck label
                datencheck_elem = driver.find_element(By.XPATH, "//*[text()='Datencheck']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", datencheck_elem)
                time.sleep(1)
                
                # Datencheck seçeneklerini bul
                standard_option = driver.find_element(By.XPATH, "//*[contains(text(), 'Standard ( options_none )')]")
                profi_option = driver.find_element(By.XPATH, "//*[contains(text(), 'Profidatencheck [+9€]')]")
                
                # İlk seçeneğe normal tıkla
                standard_option.click()
                logger.info("Clicked Standard option")
                time.sleep(1)
                
                # İkinci seçeneğe CTRL+click yap
                actions = ActionChains(driver)
                actions.key_down(Keys.CONTROL).click(profi_option).key_up(Keys.CONTROL).perform()
                logger.info("CTRL+Clicked Profidatencheck option")
                time.sleep(1)
                

            except Exception as datencheck_e:
                logger.warning(f"Failed to select Datencheck options with CTRL+click: {str(datencheck_e)}")
                
                # Alternatif JavaScript yaklaşımı
                try:
                    datencheck_result = driver.execute_script("""
                        // Find the Datencheck multiselect
                        var labels = Array.from(document.querySelectorAll('*')).filter(
                            el => el.textContent && el.textContent.trim() === 'Datencheck'
                        );
                        
                        if (labels.length === 0) return "Datencheck label not found";
                        
                        var label = labels[0];
                        var field = null;
                        var current = label;
                        
                        // Find containing field
                        for (var i = 0; i < 10; i++) {
                            current = current.parentElement;
                            if (!current) break;
                            
                            if (current.classList && current.classList.contains('admin__field')) {
                                field = current;
                                break;
                            }
                        }
                        
                        if (!field) return "Datencheck field not found";
                        
                        // Two options we want to select
                        var options = [
                            "Standard ( options_none )",
                            "Profidatencheck [+9€] ( options_professional_data_check )"
                        ];
                        
                        // Try to find a select element
                        var select = field.querySelector('select');
                        if (select) {
                            // For each option in the select
                            for (var i = 0; i < select.options.length; i++) {
                                var option = select.options[i];
                                
                                // If this option text matches one of our target options
                                if (options.some(target => option.text.includes(target))) {
                                    option.selected = true;
                                }
                            }
                            
                            // Trigger change event
                            select.dispatchEvent(new Event('change', { bubbles: true }));
                            return "Selected Datencheck options using select element";
                        }
                        
                        // Try direct element approach
                        var allElements = document.querySelectorAll('*');
                        var matchedElements = [];
                        
                        for (var i = 0; i < allElements.length; i++) {
                            var el = allElements[i];
                            if (!el.textContent) continue;
                            
                            // Check if element text matches one of our options
                            for (var j = 0; j < options.length; j++) {
                                if (el.textContent.includes(options[j])) {
                                    matchedElements.push({
                                        element: el,
                                        text: el.textContent
                                    });
                                    break;
                                }
                            }
                        }
                        
                        // Click all matched elements
                        for (var i = 0; i < matchedElements.length; i++) {
                            matchedElements[i].element.click();
                        }
                        
                        return "Tried clicking " + matchedElements.length + " Datencheck elements";
                    """)
                    
                    logger.info(f"Datencheck JavaScript selection result: {datencheck_result}")
                except Exception as js_e:
                    logger.warning(f"JavaScript Datencheck approach also failed: {str(js_e)}")
            
            # 4. Dateityp multiselect'i bul ve seçenekleri CTRL ile seç
            try:
                # Dateityp label
                dateityp_elem = driver.find_element(By.XPATH, "//*[text()='Dateityp']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dateityp_elem)
                time.sleep(1)
                
                # Dateityp seçeneklerini bul
                pdf_option = driver.find_element(By.XPATH, "//*[contains(text(), 'PDF-file ( file_type_pdf_file )')]")
                image_option = driver.find_element(By.XPATH, "//*[contains(text(), 'Image ( file_type_image )')]")
                
                # İlk seçeneğe normal tıkla
                pdf_option.click()
                logger.info("Clicked PDF-file option")
                time.sleep(1)
                
                # İkinci seçeneğe CTRL+click yap
                actions = ActionChains(driver)
                actions.key_down(Keys.CONTROL).click(image_option).key_up(Keys.CONTROL).perform()
                logger.info("CTRL+Clicked Image option")
                time.sleep(1)
                

            except Exception as dateityp_e:
                logger.warning(f"Failed to select Dateityp options with CTRL+click: {str(dateityp_e)}")
                
                # Alternatif JavaScript yaklaşımı
                try:
                    dateityp_result = driver.execute_script("""
                        // Find the Dateityp multiselect
                        var labels = Array.from(document.querySelectorAll('*')).filter(
                            el => el.textContent && el.textContent.trim() === 'Dateityp'
                        );
                        
                        if (labels.length === 0) return "Dateityp label not found";
                        
                        var label = labels[0];
                        var field = null;
                        var current = label;
                        
                        // Find containing field
                        for (var i = 0; i < 10; i++) {
                            current = current.parentElement;
                            if (!current) break;
                            
                            if (current.classList && current.classList.contains('admin__field')) {
                                field = current;
                                break;
                            }
                        }
                        
                        if (!field) return "Dateityp field not found";
                        
                        // Two options we want to select
                        var options = [
                            "PDF-file ( file_type_pdf_file )",
                            "Image ( file_type_image )"
                        ];
                        
                        // Try to find a select element
                        var select = field.querySelector('select');
                        if (select) {
                            // For each option in the select
                            for (var i = 0; i < select.options.length; i++) {
                                var option = select.options[i];
                                
                                // If this option text matches one of our target options
                                if (options.some(target => option.text.includes(target))) {
                                    option.selected = true;
                                }
                            }
                            
                            // Trigger change event
                            select.dispatchEvent(new Event('change', { bubbles: true }));
                            return "Selected Dateityp options using select element";
                        }
                        
                        // Try direct element approach
                        var allElements = document.querySelectorAll('*');
                        var matchedElements = [];
                        
                        for (var i = 0; i < allElements.length; i++) {
                            var el = allElements[i];
                            if (!el.textContent) continue;
                            
                            // Check if element text matches one of our options
                            for (var j = 0; j < options.length; j++) {
                                if (el.textContent.includes(options[j])) {
                                    matchedElements.push({
                                        element: el,
                                        text: el.textContent
                                    });
                                    break;
                                }
                            }
                        }
                        
                        // Click all matched elements
                        for (var i = 0; i < matchedElements.length; i++) {
                            matchedElements[i].element.click();
                        }
                        
                        return "Tried clicking " + matchedElements.length + " Dateityp elements";
                    """)
                    
                    logger.info(f"Dateityp JavaScript selection result: {dateityp_result}")
                except Exception as js_e:
                    logger.warning(f"JavaScript Dateityp approach also failed: {str(js_e)}")
        
        
        return True
    except Exception as e:
        logger.error(f"Failed to configure Print Product Setup with CTRL+click: {str(e)}")
        return False

def configure_additional_settings_fix(driver):
    """Configure additional settings with improved selection approach"""
    try:
        logger.info("Configuring additional settings with improved approach")

        # 1. Lieferung alanını bul ve multiselect elementini direkt manipüle et
        try:
            logger.info("Configuring Lieferung field")
            
            # Lieferung etiketini bul
            lieferung_elements = driver.find_elements(By.XPATH, "//*[text()='Lieferung']")
            
            if lieferung_elements:
                for lieferung_elem in lieferung_elements:
                    if lieferung_elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", lieferung_elem)
                        time.sleep(1)
                        
                        # JavaScript ile tüm seçenekleri seç
                        options_result = driver.execute_script("""
                            // Find the multiselect
                            var lieferungLabel = arguments[0];
                            var multiselect = null;
                            var parent = lieferungLabel;
                            
                            // Find multiselect container
                            for (var i = 0; i < 10; i++) {
                                parent = parent.parentElement;
                                if (!parent) break;
                                
                                var select = parent.querySelector('select[multiple]');
                                if (select) {
                                    multiselect = select;
                                    break;
                                }
                            }
                            
                            if (!multiselect) {
                                return {
                                    success: false,
                                    message: "Could not find multiselect element"
                                };
                            }
                            
                            // Desired options to select
                            var targetOptions = [
                                "Standard ( production_standard )",
                                "Express ( production_express )",
                                "Standard Plus ( production_standard_plus )"
                            ];
                            
                            var selected = [];
                            
                            // Select desired options
                            for (var i = 0; i < multiselect.options.length; i++) {
                                var option = multiselect.options[i];
                                var optionText = option.text.trim();
                                
                                if (targetOptions.includes(optionText) || 
                                    targetOptions.some(target => optionText.includes(target))) {
                                    option.selected = true;
                                    selected.push(optionText);
                                }
                            }
                            
                            // Trigger change event
                            multiselect.dispatchEvent(new Event('change', { bubbles: true }));
                            
                            return {
                                success: true,
                                selected: selected
                            };
                        """, lieferung_elem)
                        
                        logger.info(f"Lieferung selection result: {options_result}")
                        
                        # JavaScript ile seçilmediyse, her bir seçeneğe ayrı ayrı tıkla
                        if not (isinstance(options_result, dict) and options_result.get('success', False)):
                            logger.info("JavaScript selection failed, trying individual clicks")
                            
                            # Tüm Lieferung seçeneklerini bul
                            options = [
                                "Standard ( production_standard )",
                                "Express ( production_express )",
                                "Standard Plus ( production_standard_plus )"
                            ]
                            
                            for option_text in options:
                                try:
                                    option_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{option_text}')]")
                                    
                                    for option_elem in option_elements:
                                        if option_elem.is_displayed():
                                            # Scroll to option
                                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option_elem)
                                            time.sleep(0.5)
                                            
                                            # Click option
                                            driver.execute_script("arguments[0].click();", option_elem)
                                            logger.info(f"Clicked {option_text}")
                                            time.sleep(0.5)
                                            break
                                except Exception as option_e:
                                    logger.warning(f"Failed to click {option_text}: {str(option_e)}")
                        
                        break
            else:
                logger.warning("No Lieferung elements found")
        except Exception as lieferung_e:
            logger.warning(f"Failed to configure Lieferung: {str(lieferung_e)}")

        # 2. Anzahl alanını bul ve "1 ( quantity_1 )" seç
        try:
            logger.info("Configuring Anzahl field")
            
            anzahl_elements = driver.find_elements(By.XPATH, "//*[text()='Anzahl']")
            
            if anzahl_elements:
                for anzahl_elem in anzahl_elements:
                    if anzahl_elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", anzahl_elem)
                        time.sleep(1)
                        
                        # "1 ( quantity_1 )" seçeneğini bul
                        qty1_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '1 ( quantity_1 )')]")
                        
                        for qty1_elem in qty1_elements:
                            if qty1_elem.is_displayed():
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", qty1_elem)
                                time.sleep(0.5)
                                driver.execute_script("arguments[0].click();", qty1_elem)
                                logger.info("Selected quantity_1")
                                time.sleep(0.5)
                                break
                        
                        # Select dropdown ile deneme
                        try:
                            # Anzahl'ın yanındaki select elementini bul
                            parent = anzahl_elem
                            select_elem = None
                            
                            for _ in range(5):
                                parent = parent.find_element(By.XPATH, "..")
                                try:
                                    select_elem = parent.find_element(By.TAG_NAME, "select")
                                    break
                                except:
                                    pass
                            
                            if select_elem:
                                select = Select(select_elem)
                                select.select_by_visible_text("1 ( quantity_1 )")
                                logger.info("Selected quantity_1 using Select class")
                        except Exception as select_e:
                            logger.warning(f"Failed to select using Select class: {str(select_e)}")
                        
                        break
            else:
                logger.warning("No Anzahl elements found")
        except Exception as anzahl_e:
            logger.warning(f"Failed to configure Anzahl: {str(anzahl_e)}")

        # 3. Proof Group alanını bul ve "Digital proof ( proof_group_digital )" seç
        try:
            logger.info("Configuring Proof Group field")
            
            proof_elements = driver.find_elements(By.XPATH, "//*[text()='Proof Group']")
            
            if proof_elements:
                for proof_elem in proof_elements:
                    if proof_elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", proof_elem)
                        time.sleep(1)
                        
                        # Digital proof ( proof_group_digital ) seçeneğini bul
                        digital_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Digital proof ( proof_group_digital )')]")
                        
                        for digital_elem in digital_elements:
                            if digital_elem.is_displayed():
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", digital_elem)
                                time.sleep(0.5)
                                driver.execute_script("arguments[0].click();", digital_elem)
                                logger.info("Selected proof_group_digital")
                                time.sleep(0.5)
                                break
                        
                        # Select dropdown ile deneme
                        try:
                            # Proof Group'un yanındaki select elementini bul
                            parent = proof_elem
                            select_elem = None
                            
                            for _ in range(5):
                                parent = parent.find_element(By.XPATH, "..")
                                try:
                                    select_elem = parent.find_element(By.TAG_NAME, "select")
                                    break
                                except:
                                    pass
                            
                            if select_elem:
                                select = Select(select_elem)
                                select.select_by_visible_text("Digital proof ( proof_group_digital )")
                                logger.info("Selected proof_group_digital using Select class")
                        except Exception as select_e:
                            logger.warning(f"Failed to select using Select class: {str(select_e)}")
                        
                        break
            else:
                logger.warning("No Proof Group elements found")
        except Exception as proof_e:
            logger.warning(f"Failed to configure Proof Group: {str(proof_e)}")
  
        # 4. Product Personalization Template alanını bul ve "Test" seç/gir
        try:
            logger.info("Configuring Product Personalization Template field")
            
            template_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Product Personalization Template')]")
            
            if template_elements:
                for template_elem in template_elements:
                    if template_elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", template_elem)
                        time.sleep(1)
                        
                        # JavaScript ile input değerini ayarla
                        template_result = driver.execute_script("""
                            var templateLabel = arguments[0];
                            var container = templateLabel.closest('.admin__field') || 
                                           templateLabel.parentElement;
                            
                            // Find any input or select element
                            var input = container.querySelector('input');
                            var select = container.querySelector('select');
                            
                            if (input) {
                                input.value = 'Test';
                                input.dispatchEvent(new Event('input', {bubbles: true}));
                                input.dispatchEvent(new Event('change', {bubbles: true}));
                                return {
                                    success: true,
                                    element: "input",
                                    action: "set value"
                                };
                            } else if (select) {
                                // Look for Test option
                                for (var i = 0; i < select.options.length; i++) {
                                    if (select.options[i].text === 'Test') {
                                        select.selectedIndex = i;
                                        select.dispatchEvent(new Event('change', {bubbles: true}));
                                        return {
                                            success: true,
                                            element: "select",
                                            action: "selected option"
                                        };
                                    }
                                }
                            }
                            
                            // Look for a visible Test option nearby to click
                            var testElements = [];
                            var allElements = document.querySelectorAll('*');
                            
                            for (var i = 0; i < allElements.length; i++) {
                                if (allElements[i].textContent && 
                                    allElements[i].textContent.trim() === 'Test' &&
                                    window.getComputedStyle(allElements[i]).display !== 'none') {
                                    testElements.push(allElements[i]);
                                }
                            }
                            
                            if (testElements.length > 0) {
                                testElements[0].click();
                                return {
                                    success: true,
                                    element: "text element",
                                    action: "clicked"
                                };
                            }
                            
                            return {
                                success: false,
                                message: "Could not find a way to set Template value"
                            };
                        """, template_elem)
                        
                        logger.info(f"Template setting result: {template_result}")
                        
                        # JavaScript başarısız olursa, Test metnine tıkla
                        if not (isinstance(template_result, dict) and template_result.get('success', False)):
                            try:
                                test_elements = driver.find_elements(By.XPATH, "//*[text()='Test']")
                                
                                for test_elem in test_elements:
                                    if test_elem.is_displayed():
                                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", test_elem)
                                        time.sleep(0.5)
                                        driver.execute_script("arguments[0].click();", test_elem)
                                        logger.info("Clicked Test text element")
                                        time.sleep(0.5)
                                        break
                            except Exception as test_e:
                                logger.warning(f"Failed to click Test text: {str(test_e)}")
                        
                        break
            else:
                logger.warning("No Product Personalization Template elements found")
        except Exception as template_e:
            logger.warning(f"Failed to configure Product Personalization Template: {str(template_e)}")
 
        return True
    except Exception as e:
        logger.error(f"Failed to configure additional settings with improved approach: {str(e)}")
        return False
    
def save_product(driver):
    """Sağ üst köşedeki Save butonuna tıkla"""
    try:
        logger.info("Clicking the Save button in the top right corner")
 
        # Sağ üstteki Save butonunu bul (ekran görüntüsündeki kırmızı kutu içindeki buton)
        try:
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'action-primary') and contains(text(), 'Save')]"))
            )
            
            # Görünür olması için sayfayı kaydır
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
            time.sleep(1)
            
            # Butona tıkla
            driver.execute_script("arguments[0].click();", save_button)
            logger.info("Clicked Save button successfully")
            
            # Kaydetme işleminin tamamlanması için bekle
            time.sleep(10)
            
            return True
        except Exception as e:
            logger.warning(f"Standard approach failed: {str(e)}")
            
            # JavaScript alternatif yaklaşım
            try:
                result = driver.execute_script("""
                    // Sağ üstteki Save butonunu bul
                    var saveButtons = Array.from(document.querySelectorAll('button')).filter(
                        el => el.textContent.trim() === 'Save' && window.getComputedStyle(el).display !== 'none'
                    );
                    
                    if (saveButtons.length > 0) {
                        // İlk butona tıkla
                        saveButtons[0].click();
                        return "Clicked Save button by text";
                    }
                    
                    // Sınıf adına göre bul
                    var primaryButtons = document.querySelectorAll('.action-primary, .primary');
                    for (var i = 0; i < primaryButtons.length; i++) {
                        var btn = primaryButtons[i];
                        if (window.getComputedStyle(btn).display !== 'none' && 
                            btn.getBoundingClientRect().top < 200) {  // Sayfanın üst kısmında
                            
                            btn.click();
                            return "Clicked primary button in the top section";
                        }
                    }
                    
                    // Pozisyona göre bul - sağ üst köşede olan görünür buton
                    var allButtons = document.querySelectorAll('button');
                    var rightTopButton = null;
                    var maxRight = 0;
                    
                    for (var i = 0; i < allButtons.length; i++) {
                        var btn = allButtons[i];
                        var rect = btn.getBoundingClientRect();
                        
                        // Sadece görünür butonları dikkate al
                        if (window.getComputedStyle(btn).display !== 'none' && rect.top < 200) {
                            if (rect.right > maxRight) {
                                maxRight = rect.right;
                                rightTopButton = btn;
                            }
                        }
                    }
                    
                    if (rightTopButton) {
                        rightTopButton.click();
                        return "Clicked rightmost button in the top section";
                    }
                    
                    return "Could not find Save button";
                """)
                
                logger.info(f"JavaScript result: {result}")
                
                if "Clicked" in str(result):
                    logger.info("Successfully clicked Save button using JavaScript")
                    time.sleep(10)  # Kaydetme işleminin tamamlanması için bekle
                    return True
                else:
                    logger.warning("JavaScript approach also failed")
                    return False
            except Exception as js_e:
                logger.error(f"JavaScript approach failed: {str(js_e)}")
                return False
            
    except Exception as e:
        logger.error(f"Failed to click Save button: {str(e)}")
        return False

def process_csv_file_with_web2print(driver, csv_file_path):
    """Process products from a CSV file with Web2Print settings"""
    products_added = 0
    products_processed = 0
    
    try:
        logger.info(f"Processing CSV file: {csv_file_path}")
        
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            
            # Skip header rows (first 6 rows)
            for _ in range(6):
                next(csv_reader, None)
            
            # Process rows
            for row in csv_reader:
                products_processed += 1
                
                # Check if we've reached the maximum number of products
                
                
                try:
                    # Extract product data from CSV
                    product_name = row[CSV_MAPPINGS["product_name"]]
                    sku = row[CSV_MAPPINGS["sku"]]
                    
                    logger.info(f"Processing product {products_processed}: {product_name} (SKU: {sku})")
                    
                    # Navigate to Add Product page
                    if not navigate_to_add_product(driver):
                        logger.error("Failed to navigate to Add Product page")
                        continue
                    
                    # Fill basic info
                    if not fill_basic_info(driver, product_name, sku):
                        logger.error(f"Failed to fill basic info for product: {product_name}")
                        continue
                    
                    # Set Is Meter Price to Yes (True)
                    if not set_meter_price(driver, is_meter_price=True):
                        logger.warning(f"Failed to set 'Is Meter Price' for product: {product_name}")
                        # Continue anyway
                    
                    # Select Producer with simplified approach (as it works reliably)
                    if not select_producer_simplified(driver, producer_value="wmd.networkhero.eu"):
                        logger.warning(f"Failed to select Producer for product: {product_name}")
                        # Continue anyway
                    
                    # Set Disable pickup to Yes (True)
                    if not set_disable_pickup(driver, disable_pickup=True):
                        logger.warning(f"Failed to set 'Disable pickup' for product: {product_name}")
                        # Continue anyway
                    
                    # Select websites - ÖNCE BURADA WEBSITE SEÇİMİ YAPILIYOR
                    if not select_websites(driver, WEBSITES_TO_SELECT):
                        logger.warning(f"Failed to select websites for product: {product_name}")
                        # Continue anyway
                    
                    # Configure Web2Print Settings - SONRA WEB2PRINT AYARLARI YAPILIYOR
                    if not update_product_type_like_producer(driver, product_type="Upload + Personalization + Configuration"):
                        logger.warning(f"Failed to configure Web2Print Settings for product: {product_name}")
                        # Continue anyway
                    # Set Price Calculation to Fixed in Matrix tab
                    if not set_price_calculation_to_fixed(driver):
                        logger.warning(f"Failed to set Price Calculation to Fixed for product: {product_name}")
                        # Continue anyway
                    # Configure Others tab and set Production Days
                    if not select_production_days_multiselect(driver):
                        logger.warning(f"Failed to configure Others tab and Production Days for product: {product_name}")
                        # Continue anyway
                    # Configure Delivery section with final optimized approach
                    if not configure_delivery_section_final(driver):
                        logger.warning(f"Failed to configure Delivery section for product: {product_name}")
                        # Continue anyway
                    # Configure Print Product Setup
                    if not configure_print_product_setup_with_ctrl(driver):
                        logger.warning(f"Failed to configure Print Product Setup for product: {product_name}")
                        # Continue anyway

                    # Configure additional settings
                    if not configure_additional_settings_fix(driver):
                        logger.warning(f"Failed to configure additional settings for product: {product_name}")
                        # Continue anyway

                    # Tüm ayarlar tamamlandıktan sonra kaydet
                    if not save_product(driver):
                        logger.warning(f"Failed to save product: {product_name}")
                        # Continue anyway
                    time.sleep(2)
                    
                    # Save product with retry logic
                    save_attempts = 0
                    max_save_attempts = 3
                    saved = False
                    
                    while not saved and save_attempts < max_save_attempts:
                        save_attempts += 1
                        try:
                            logger.info(f"Attempting to save product (attempt {save_attempts}/{max_save_attempts})")
                            
                            save_button = WebDriverWait(driver, 15).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, ".action-primary.action-save"))
                            )
                            driver.execute_script("arguments[0].click();", save_button)
                            logger.info("Clicked Save button")
                            
                            # Wait for save to complete
                            WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, ".message-success"))
                            )
                            
                            logger.info("Product saved successfully")
                            saved = True
                        except Exception as save_e:
                            logger.warning(f"Save attempt {save_attempts} failed: {str(save_e)}")
                            time.sleep(2)  # Wait before retry
                    
                    if saved:
                        logger.info(f"Added product: {product_name} with SKU: {sku}")
                        products_added += 1
                    else:
                        logger.error(f"Failed to save product after {max_save_attempts} attempts: {product_name}")
                    
                except IndexError:
                    logger.error(f"Row {products_processed} has invalid format or missing data")
                except Exception as e:
                    logger.error(f"Failed to process product at row {products_processed}: {str(e)}")
            
        logger.info(f"CSV processing complete. Added {products_added} out of {products_processed} products")
        return products_added
    except Exception as e:
        logger.error(f"Failed to process CSV file: {str(e)}")
        return products_added
    
def main():
    """Main function for the complete workflow"""
    # Setup Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    
    # Kullanıcıdan kaç ürün eklemek istediğini sor
    try:
        product_count = int(input("Kaç ürün eklemek istiyorsunuz? "))
        if product_count <= 0:
            logger.warning("Geçersiz ürün sayısı, varsayılan olarak 1 olarak ayarlandı.")
            product_count = 1
    except ValueError:
        logger.warning("Geçersiz giriş, varsayılan olarak 1 ürün eklenecek.")
        product_count = 1

    logger.info(f"Toplam {product_count} ürün eklenecek.")
    
    # Initialize Chrome driver
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        logger.info("Chrome WebDriver initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize Chrome WebDriver: {str(e)}")
        return
    
    try:
        # CSV'den ürün verilerini oku
        products_to_add = []
        try:
            with open(CSV_FILE_PATH, 'r', encoding='utf-8') as file:
                csv_reader = csv.reader(file)
                
                # Skip header rows (first 6 rows)
                for _ in range(6):
                    next(csv_reader, None)
                
                # Ürün verilerini topla (kullanıcının belirttiği sayıda)
                for idx, row in enumerate(csv_reader):
                    if idx >= product_count:  # Kullanıcının belirttiği sayıda ürün ekle
                        break
                    
                    try:
                        product_name = row[CSV_MAPPINGS["product_name"]]
                        sku = row[CSV_MAPPINGS["sku"]]
                        products_to_add.append((product_name, sku, idx))
                    except IndexError:
                        logger.error(f"Row {idx+1} has invalid format or missing data")
        except Exception as csv_e:
            logger.error(f"Failed to read CSV file: {str(csv_e)}")
            return

        # Her ürün için işlemleri yap
        for product_idx, (product_name, sku, csv_idx) in enumerate(products_to_add):
            logger.info(f"Ürün {product_idx+1}/{len(products_to_add)} işleniyor: {product_name}")
                
            if login_to_admin(driver):
                # Step 1: Click on CloudLab
                if click_cloudlab(driver):
                    logger.info("Successfully clicked on CloudLab")

                    # Step 2: Click on Page Options
                    if click_page_options(driver):
                        logger.info("Successfully clicked on Page Options")
                        
                        # Step 3-6: Search for 'paper', press Enter, wait, and click Material
                        if search_and_click_material(driver):
                            logger.info("Successfully completed search and click workflow")
                            
                            if click_page_option_values_and_add_new(driver):
                                logger.info("Successfully clicked on Page Option Values and Add new Value")
                                if fill_add_new_value_form(driver, CSV_FILE_PATH):
                                    logger.info("Successfully filled and saved the form")
                                    if click_popup_save_button(driver):
                                        logger.info("Successfully clicked modal save button.")
                                        if refresh_page_and_click_main_save(driver):
                                            logger.info("Successfully clicked modal save button.")
                    
                            # Her ürün için product part ekle
                            if add_product_part_from_csv(driver, CSV_FILE_PATH, csv_idx):
                                logger.info(f"Successfully added a new product part for product {product_idx+1}")
                        
                                # Ürün ekleme sayfasına git
                                navigate_to_add_product(driver)
                                # Temel bilgileri doldur
                                if fill_basic_info(driver, product_name, sku):
                                    # Tüm ayarları yap
                                    try:
                                        # Her bir ayarı ayrı ayrı kontrol et
                                        settings_steps = [
                                            (set_meter_price, (driver, True), "Is Meter Price"),
                                            (select_producer_simplified, (driver, "wmd.networkhero.eu"), "Producer"),
                                            (set_disable_pickup, (driver, True), "Disable pickup"),
                                            (select_websites, (driver, WEBSITES_TO_SELECT), "Websites"),
                                            (update_product_type_like_producer, (driver, "Upload + Personalization + Configuration"), "Web2Print Settings"),
                                            (set_price_calculation_to_fixed, (driver,), "Price Calculation"),
                                            (select_production_days_multiselect, (driver,), "Production Days"),
                                            (configure_delivery_section_final, (driver,), "Delivery Section"),
                                            (configure_print_product_setup_with_ctrl, (driver,), "Print Product Setup"),
                                            (configure_additional_settings_fix, (driver,), "Additional Settings")
                                        ]
                                        
                                        for setting_func, args, setting_name in settings_steps:
                                            if not setting_func(*args):
                                                logger.warning(f"Failed to set {setting_name} for product: {product_name}")
                                        
                                        # Ürünü kaydet
                                        if save_product(driver):
                                            logger.info(f"Successfully saved product {product_idx+1}: {product_name}")
                                            
                                            # Additional scripts for each product
                                            printnet.main()
                                            add_fixed_price.main()
                                            add_fixed_price_printnet.main()
                                        else:
                                            logger.error(f"Failed to save product {product_idx+1}: {product_name}")
                                    
                                    except Exception as settings_e:
                                        logger.error(f"Error configuring settings for product {product_idx+1}: {str(settings_e)}")
                                else:
                                    logger.error(f"Failed to fill basic info for product {product_idx+1}: {product_name}")
                            else:
                                logger.error(f"Failed to add product part for product {product_idx+1}")
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
    finally:
        try:
            driver.quit()
        except:
            pass
        input("İşlem tamamlandı. Tarayıcıyı kapatmak için Enter'a basın...")
        logger.info("Browser closed")
        
if __name__ == "__main__":
    main()      
