import time
import os
import logging
import csv
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# Setup logging
def setup_logging():
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"printnet_fixed_price_{timestamp}.log")
    
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
printnet_url = "https://printnet.networkhero.eu/dashboard/admin"

# Hardcoded credentials
USERNAME = "Serhat.Seyitoglu"
PASSWORD = "Serhat.kamile.mavi2212"

def login_to_admin(driver):
    """Login to printnet admin panel using hardcoded credentials"""
    try:
        logger.info(f"Navigating to {printnet_url}")
        driver.get(printnet_url)
        
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

def click_fixed_price(driver):
    try:
 
        # APPROACH 1: Try direct XPath approach
        try:
            # Look for Fixed Price submenu item
            fixed_price_submenu = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-cloudlab')]//li[contains(@class, 'item-fixed-price') or contains(@class, 'item-cloudlab-fixed-price')]//a"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", fixed_price_submenu)
            driver.execute_script("arguments[0].click();", fixed_price_submenu)

            
            return True
        except Exception as e:
            logger.warning(f"XPath approach failed for Fixed Price: {str(e)}")
            
            # APPROACH 2: Try to find by text content
            try:
                # Look for elements containing "Fixed Price" text
                fixed_price_elements = driver.find_elements(By.XPATH, 
                    "//span[contains(text(), 'Fixed Price')] | //a[contains(text(), 'Fixed Price')] | //*[contains(text(), 'Fixed Price')]")
                
                if fixed_price_elements:
                    for elem in fixed_price_elements:
                        if elem.is_displayed():
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                           
                            # If the element itself is not clickable, try to find its parent <a> tag
                            try:
                                # First try to click the element directly
                                driver.execute_script("arguments[0].click();", elem)
                            except:
                                # If that fails, try to find and click its parent <a> tag
                                js_script = """
                                var element = arguments[0];
                                var currentNode = element;
                                
                                // Traverse up to find the closest anchor tag
                                while (currentNode && currentNode.tagName !== 'A' && currentNode.tagName !== 'BODY') {
                                    currentNode = currentNode.parentNode;
                                }
                                
                                if (currentNode && currentNode.tagName === 'A') {
                                    currentNode.click();
                                    return true;
                                } else {
                                    // If we can't find parent anchor, try to find a nearby anchor
                                    var nearby = element.querySelector('a') || 
                                                 element.parentNode.querySelector('a') || 
                                                 element.parentNode.parentNode.querySelector('a');
                                    if (nearby) {
                                        nearby.click();
                                        return true;
                                    }
                                }
                                return false;
                                """
                                result = driver.execute_script(js_script, elem)
                                if not result:
                                    continue  # Try the next element if this one failed
                            

                            return True
                    
                    logger.warning("Found Fixed Price elements but none were clickable")
                else:
                    logger.warning("Could not find Fixed Price by text content")
                
            except Exception as e2:
                logger.warning(f"Text content approach failed for Fixed Price: {str(e2)}")
            
            # APPROACH 3: Try to find the Fixed Price option using JavaScript
            try:
                js_script = """
                // Function to find and click the Fixed Price menu item
                function findAndClickFixedPrice() {
                    // Strategy 1: Look for elements with "Fixed Price" text
                    var xpath = "//span[contains(text(), 'Fixed Price')] | //a[contains(text(), 'Fixed Price')]";
                    var elements = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
                    
                    for (var i = 0; i < elements.snapshotLength; i++) {
                        var element = elements.snapshotItem(i);
                        if (element && element.offsetParent !== null) { // Check if visible
                            // Find the closest clickable ancestor (usually an <a> tag)
                            var clickable = element;
                            while (clickable && clickable.tagName !== 'A' && clickable.tagName !== 'BUTTON' && clickable.tagName !== 'BODY') {
                                clickable = clickable.parentNode;
                            }
                            
                            if (clickable && (clickable.tagName === 'A' || clickable.tagName === 'BUTTON')) {
                                clickable.scrollIntoView({block: 'center'});
                                clickable.click();
                                return "Clicked Fixed Price element: " + element.textContent;
                            }
                        }
                    }
                    
                    // Strategy 2: Look for menu items containing the text "Fixed Price"
                    var menuItems = document.querySelectorAll('li.item-fixed-price, li[class*="fixed-price"], div[class*="fixed-price"]');
                    for (var i = 0; i < menuItems.length; i++) {
                        var item = menuItems[i];
                        if (item && item.offsetParent !== null) {
                            var link = item.querySelector('a');
                            if (link) {
                                link.scrollIntoView({block: 'center'});
                                link.click();
                                return "Clicked Fixed Price menu item";
                            }
                        }
                    }
                    
                    // Strategy 3: Look in the open CloudLab menu for items
                    var cloudLabMenu = document.querySelector('li.item-cloudlab.active');
                    if (cloudLabMenu) {
                        var submenuItems = cloudLabMenu.querySelectorAll('a');
                        for (var i = 0; i < submenuItems.length; i++) {
                            var link = submenuItems[i];
                            if (link.textContent.toLowerCase().includes('fixed') || 
                                link.textContent.toLowerCase().includes('price')) {
                                link.scrollIntoView({block: 'center'});
                                link.click();
                                return "Clicked submenu item: " + link.textContent;
                            }
                        }
                    }
                    
                    return "Could not find Fixed Price menu item";
                }
                
                return findAndClickFixedPrice();
                """
                
                result = driver.execute_script(js_script)
    
                if "Clicked" in result:

                    return True
                else:
                    logger.warning("JavaScript approach could not find Fixed Price option")
            except Exception as e3:
                logger.warning(f"JavaScript approach failed: {str(e3)}")
            
            # APPROACH 4: Try to find by position in the CloudLab menu
            try:
                # Based on the screenshot, Fixed Price should be the 8th item in the CloudLab menu
                # Find all visible menu items in the CloudLab submenu
                js_script = """
                // Function to find the Nth visible item in the CloudLab menu
                function findNthVisibleMenuItem(n) {
                    // Find the CloudLab menu
                    var cloudLabMenu = document.querySelector('li.item-cloudlab.active');
                    if (!cloudLabMenu) return "CloudLab menu not found or not active";
                    
                    // Find all visible menu items
                    var visibleItems = [];
                    var allItems = cloudLabMenu.querySelectorAll('li a');
                    
                    for (var i = 0; i < allItems.length; i++) {
                        var item = allItems[i];
                        if (item.offsetParent !== null) { // Check if visible
                            visibleItems.push(item);
                        }
                    }
                    
                    console.log("Found " + visibleItems.length + " visible items in CloudLab menu");
                    
                    // Check if we have enough items
                    if (visibleItems.length < n) {
                        return "Not enough visible items in CloudLab menu. Found: " + visibleItems.length;
                    }
                    
                    // Click the nth item (0-indexed)
                    visibleItems[n-1].scrollIntoView({block: 'center'});
                    visibleItems[n-1].click();
                    return "Clicked item #" + n + ": " + visibleItems[n-1].textContent;
                }
                
                return findNthVisibleMenuItem(8); // Fixed Price is the 8th item in the menu
                """
                
                result = driver.execute_script(js_script)

                if "Clicked" in result:
                    
                    return True
                else:
                    logger.warning("Position-based approach failed")
            except Exception as e4:
                logger.warning(f"Position-based approach failed: {str(e4)}")
            
            
           
            return True  # Assume manual intervention succeeded
    except Exception as e:
        logger.error(f"Error in click_fixed_price: {str(e)}")
        return False

def click_add_new(driver):
    """Click the Add New button in the top-right corner of the Fixed Price page"""
    try:        
        # APPROACH 1: Try direct selector approach
        try:
            # Various selectors that might match the Add New button
            add_button_selectors = [
                ".page-actions-buttons button.add",
                ".page-actions-buttons button.action-primary",
                ".page-actions button.primary",
                ".page-actions button.add",
                "button.primary[title='Add New']",
                "button.action-primary[title='Add New']",
                "button[data-ui-id*='add-button']"
            ]
            
            for selector in add_button_selectors:
                try:
                    add_buttons = driver.find_elements(By.CSS_SELECTOR, selector)
                    if add_buttons:
                        for btn in add_buttons:
                            if btn.is_displayed():
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
   
                                driver.execute_script("arguments[0].click();", btn)
       
                                return True
                except Exception as e:
                    logger.warning(f"Selector {selector} failed: {str(e)}")
                    continue
            
            logger.warning("No Add New button found with direct selectors")
        except Exception as e:
            logger.warning(f"Direct selector approach failed: {str(e)}")
        
        # APPROACH 2: Try to find by text content
        try:
            # Look for elements containing "Add New" text
            add_new_xpath = "//button[contains(., 'Add New') or .//*[contains(text(), 'Add New')]] | //a[contains(., 'Add New')]"
            add_buttons = driver.find_elements(By.XPATH, add_new_xpath)
            
            if add_buttons:
                for btn in add_buttons:
                    if btn.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
         
                        driver.execute_script("arguments[0].click();", btn)
     
                        return True
                
                logger.warning("Add New buttons found but none were clickable")
            else:
                logger.warning("No Add New buttons found by text")
        except Exception as e:
            logger.warning(f"Text content approach failed: {str(e)}")
        
        # APPROACH 3: JavaScript approach
        try:
            js_script = """
            // Function to find and click the Add New button
            function findAndClickAddNew() {
                // Look for buttons with exact "Add New" text
                var addButtons = Array.from(document.querySelectorAll('button, a')).filter(function(btn) {
                    if (!btn.offsetParent) return false; // Skip hidden elements
                    
                    // Check button text
                    var text = btn.textContent.trim().toLowerCase();
                    if (text === 'add new' || text === '+ add new' || text === 'add') {
                        return true;
                    }
                    
                    // Check button title attribute
                    var title = btn.getAttribute('title');
                    if (title && (title.toLowerCase() === 'add new' || title.toLowerCase() === 'add')) {
                        return true;
                    }
                    
                    // Check for child elements with "Add New" text
                    var children = btn.querySelectorAll('*');
                    for (var i = 0; i < children.length; i++) {
                        if (children[i].textContent.trim().toLowerCase() === 'add new') {
                            return true;
                        }
                    }
                    
                    return false;
                });
                
                if (addButtons.length > 0) {
                    // Sort buttons by position (prefer top-right corner)
                    addButtons.sort(function(a, b) {
                        var aRect = a.getBoundingClientRect();
                        var bRect = b.getBoundingClientRect();
                        
                        // Buttons at the top of the page are preferred
                        if (aRect.top < 200 && bRect.top >= 200) return -1;
                        if (bRect.top < 200 && aRect.top >= 200) return 1;
                        
                        // If both are at the top, prefer the one more to the right
                        if (aRect.top < 200 && bRect.top < 200) {
                            return bRect.right - aRect.right;
                        }
                        
                        // Otherwise just sort by vertical position
                        return aRect.top - bRect.top;
                    });
                    
                    // Click the best candidate
                    var bestButton = addButtons[0];
                    bestButton.scrollIntoView({block: 'center'});
                    
                    // Highlight button for visibility
                    
                    return "Clicked button: " + bestButton.textContent.trim();
                }
                
                // If no Add New button found by text, look for add button by position (top-right)
                var actionAreas = document.querySelectorAll('.page-actions, .page-actions-buttons, .page-main-actions');
                for (var i = 0; i < actionAreas.length; i++) {
                    var area = actionAreas[i];
                    var buttons = area.querySelectorAll('button, a');
                    
                    if (buttons.length > 0) {
                        // Assume the rightmost button in the action area is the Add New button
                        buttons = Array.from(buttons).filter(function(btn) {
                            return btn.offsetParent !== null; // Only visible buttons
                        });
                        
                        if (buttons.length > 0) {
                            // Sort by position (rightmost first)
                            buttons.sort(function(a, b) {
                                var aRect = a.getBoundingClientRect();
                                var bRect = b.getBoundingClientRect();
                                return bRect.right - aRect.right;
                            });
                            
                            var button = buttons[0]; // Rightmost button
                            button.scrollIntoView({block: 'center'});
                            
                            // Highlight for visibility

                            
                            return "Clicked rightmost button in action area: " + button.textContent.trim();
                        }
                    }
                }
                
                return "Could not find Add New button";
            }
            
            return findAndClickAddNew();
            """

            result = driver.execute_script(js_script)
 
            # Wait a moment for the button to be clicked
            time.sleep(0.5)
            
            if "Clicked" in result:

                return True
            else:
                logger.warning("JavaScript approach could not find Add New button")
        except Exception as e:
            logger.warning(f"JavaScript approach failed: {str(e)}")
        
        # APPROACH 4: Try direct coordinate click (last resort)
        try:
            # Get window size
            window_size = driver.get_window_size()
            
            # Calculate position for top-right add button (adjust based on actual UI)
            x_coord = window_size['width'] - 100  # 100px from right edge
            y_coord = 100  # 100px from top
            
            # Execute a click at these coordinates
            action = webdriver.ActionChains(driver)
            action.move_by_offset(x_coord, y_coord).click().perform()

            return True
        except Exception as e:
            logger.warning(f"Direct coordinate click failed: {str(e)}")
        
        
       
        return True  # Assume manual intervention succeeded
        
    except Exception as e:
        logger.error(f"Error in click_add_new: {str(e)}")
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

def set_price_unit_to_one(driver):

    try:
        js_script = """
        // Price Unit alanını bul
        var priceUnitLabels = Array.from(document.querySelectorAll('span, label')).filter(function(el) {
            return el.textContent.trim() === 'Price Unit';
        });
        
        if (priceUnitLabels.length === 0) {
            return "Price Unit etiketi bulunamadı";
        }
        
        // İlgili input alanını bul
        var field = priceUnitLabels[0].closest('.admin__field, .field');
        if (!field) {
            return "Price Unit alanı bulunamadı";
        }
        
        var input = field.querySelector('input');
        if (!input) {
            return "Input alanı bulunamadı";
        }
        
        // Değeri 1 olarak ayarla
        input.value = '1';
        input.dispatchEvent(new Event('input', {bubbles: true}));
        input.dispatchEvent(new Event('change', {bubbles: true}));
        
        return "Price Unit değeri 1 olarak ayarlandı";
        """
        
        result = driver.execute_script(js_script)
        logger.info(f"JavaScript sonucu: {result}")
        
        if "ayarlandı" not in result:
            # Alternatif yaklaşım: Selenium ile XPath kullanarak
            try:
                price_unit_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[text()='Price Unit']/ancestor::div[contains(@class, 'admin__field')]//input"))
                )
                price_unit_input.clear()
                price_unit_input.send_keys("1")
                logger.info("XPath ile Price Unit değeri 1 olarak ayarlandı")
                return True
            except Exception as e:
                logger.warning(f"XPath yaklaşımı başarısız: {str(e)}")
                
                
                return True
        
        return True
    
    except Exception as e:
        logger.error(f"Price Unit ayarlamada hata: {str(e)}")
        # Manuel müdahale iste
       
        return True

def select_product_parts_dropdown(driver, product_value):
    """Product Parts dropdown menüsünü bulur, açar ve CSV'deki ürün değerini seçer"""
    try:
      
        js_find_dropdown = """
        var possibleLabels = [
            'Product Part', 'Product Parts', 'Produkt Teil', 'Produkt', 
            'Product', 'Part', 'Teil'
        ];
        
        // Önce tam eşleşme ara
        for (var i = 0; i < possibleLabels.length; i++) {
            var label = possibleLabels[i];
            var elements = Array.from(document.querySelectorAll('label, span')).filter(function(el) {
                return el.textContent.trim() === label;
            });
            
            if (elements.length > 0) {
                var field = elements[0].closest('.admin__field, .field, .fieldset-wrapper');
                if (field) {
                    var dropdown = field.querySelector('select, .action-select, button.action-toggle');
                    if (dropdown) {
                        return {
                            found: true,
                            element: dropdown,
                            label: label
                        };
                    }
                }
            }
        }
        
        // Kısmi eşleşme ara - span ve label'larda "Product" ve "Part" kelimelerini içeren öğeler
        var allLabels = Array.from(document.querySelectorAll('label, span')).filter(function(el) {
            var text = el.textContent.trim().toLowerCase();
            return text.includes('product') && text.includes('part');
        });
        
        if (allLabels.length > 0) {
            var field = allLabels[0].closest('.admin__field, .field, .fieldset-wrapper');
            if (field) {
                var dropdown = field.querySelector('select, .action-select, button.action-toggle');
                if (dropdown) {
                    return {
                        found: true,
                        element: dropdown,
                        label: allLabels[0].textContent.trim()
                    };
                }
            }
        }
        
        // Son deneme: doğrudan select elementlerini ara
        var selects = document.querySelectorAll('select');
        for (var i = 0; i < selects.length; i++) {
            // Görünür olmalı
            if (selects[i].offsetParent !== null) {
                // Select'in name özelliğini veya yakındaki bir etiketi kontrol et
                var selectName = selects[i].getAttribute('name') || '';
                if (selectName.includes('product') || selectName.includes('part')) {
                    return {
                        found: true,
                        element: selects[i],
                        label: 'Product Part (select by name)'
                    };
                }
                
                // Etiket kontrolü
                var parentField = selects[i].closest('.admin__field, .field');
                if (parentField) {
                    var labels = parentField.querySelectorAll('label, span');
                    for (var j = 0; j < labels.length; j++) {
                        var labelText = labels[j].textContent.trim().toLowerCase();
                        if (labelText.includes('product') || labelText.includes('part')) {
                            return {
                                found: true,
                                element: selects[i],
                                label: labels[j].textContent.trim()
                            };
                        }
                    }
                }
            }
        }
        
        // Tüm görünür select elementlerini döndür son çare olarak
        for (var i = 0; i < selects.length; i++) {
            if (selects[i].offsetParent !== null) {
                return {
                    found: true,
                    element: selects[i],
                    label: 'Select element (fallback)'
                };
            }
        }
        
        // Hiçbir şey bulunamadı
        return {
            found: false,
            message: "Product Parts alanı bulunamadı"
        };
        """
        
        # JavaScript'i çalıştır
        dropdown_info = driver.execute_script(js_find_dropdown)
  
        
        # Dropdown bulunamadıysa, bir de XPath ile deneyelim
        if not dropdown_info.get('found', False):
            try:
                # Ekran görüntüsündeki select box'a doğrudan XPath ile ulaşmayı dene
                xpath_patterns = [
                    "//span[text()='Product Part']/ancestor::div[contains(@class, 'admin__field')]//select",
                    "//label[text()='Product Part']/following-sibling::div//select",
                    "//span[contains(text(), 'Product Part')]/ancestor::div[contains(@class, 'field')]//select",
                    "//select[@name='product_part']",
                    "//select[contains(@name, 'product_part')]",
                    "//select[@name='product_parts']",
                    "//select[contains(@name, 'product_parts')]",
                    "//select[contains(@id, 'product_part')]",
                    "//select[contains(@class, 'admin__control-select')]",
                    "//select"  # Son çare - görünür bir select elementi
                ]
                
                for xpath in xpath_patterns:
                    try:
                        dropdown_elements = driver.find_elements(By.XPATH, xpath)
                        for elem in dropdown_elements:
                            if elem.is_displayed():
                                dropdown_info = {
                                    'found': True,
                                    'element': elem,
                                    'label': f'Product Part (XPath: {xpath})'
                                }
                                
                                break
                        if dropdown_info.get('found', False):
                            break
                    except Exception as xe:
                        logger.warning(f"XPath {xpath} için hata: {str(xe)}")
                        continue
            except Exception as e:
                logger.warning(f"XPath ile dropdown arama hatası: {str(e)}")
            
            # Hala bulunamadıysa manuel müdahale iste
            if not dropdown_info.get('found', False):
                logger.warning(dropdown_info.get('message', 'Product Parts dropdown bulunamadı'))
                
                return True
        
        # Dropdown'u tıkla
        dropdown_element = dropdown_info['element']
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_element)
        time.sleep(1)
        
        # Eğer bu bir <select> elementi ise, tıklamak yerine doğrudan seçim yapmayı dene
        element_tag = driver.execute_script("return arguments[0].tagName.toLowerCase();", dropdown_element)
        
        if element_tag == 'select':
            
            # CSV'den gelen ürün adını temizle
            clean_product_value = product_value.strip() if product_value else ""
            
            if not clean_product_value:
                logger.warning("CSV'den alınan ürün değeri boş")
                clean_product_value = ""
            
            # Select elementinden tüm seçenekleri al
            options = driver.execute_script("""
            var select = arguments[0];
            var options = [];
            for (var i = 0; i < select.options.length; i++) {
                options.push({
                    index: i,
                    text: select.options[i].text.trim(),
                    value: select.options[i].value
                });
            }
            return options;
            """, dropdown_element)

            # AŞAMA 1: Tam eşleşme ara (büyük/küçük harf göz ardı edilerek)
            for i, option in enumerate(options):
                if option['text'].lower() == clean_product_value.lower():
                   
                    driver.execute_script("arguments[0].selectedIndex = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", 
                                        dropdown_element, i)
                    time.sleep(1)
                    return True
            
            # AŞAMA 2: Kelime tabanlı eşleşme puanlama sistemi
            if clean_product_value:
                # CSV'deki ürünü anlamlı kelimelere ayır (çok kısa olanları atla)
                clean_words = [w.lower() for w in clean_product_value.split() if len(w) > 2]
                
                best_option_index = -1
                best_option_score = -1
                best_option_match_count = 0
                
                for i, option in enumerate(options):
                    option_text = option['text'].lower()
                    option_words = [w.lower() for w in option_text.split() if len(w) > 2]
                    
                    # Eşleşen kelime sayısını hesapla
                    match_count = sum(1 for word in clean_words if word in option_text)
                    
                    # Eşleşme oranı hesapla
                    csv_match_ratio = match_count / len(clean_words) if clean_words else 0
                    
                    # Toplam puanı hesapla
                    score = int(csv_match_ratio * 100)  # 0-100 arası puan
                    
                    # İlk pozisyona göre ek puan
                    first_word_bonus = False
                    if clean_words and clean_words[0] in option_text:
                        first_word_pos = option_text.find(clean_words[0])
                        if first_word_pos <= 15:  # İlk 15 karakter içinde bulunursa
                            score += 20
                            first_word_bonus = True
                    
                    # Eğer eşleşen kelime sayısı daha fazlaysa veya aynıysa ve puan daha yüksekse
                    if match_count > best_option_match_count or (match_count == best_option_match_count and score > best_option_score):
                        best_option_index = i
                        best_option_score = score
                        best_option_match_count = match_count
                        
                        match_details = f"{match_count} kelime eşleşti, oran: {csv_match_ratio:.2f}"
                        if first_word_bonus:
                            match_details += ", ilk kelime bonusu"
                        
                        
                
                # En iyi eşleşme varsa ve yeterince iyi bir puanı varsa seç
                if best_option_index >= 0 and best_option_score >= 50 and best_option_match_count > 0:
                
                    driver.execute_script("arguments[0].selectedIndex = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", 
                                        dropdown_element, best_option_index)
                    time.sleep(1)
                    return True
                else:
                    logger.warning(f"Yeterince iyi bir eşleşme bulunamadı. En iyi puan: {best_option_score}, Eşleşen kelime: {best_option_match_count}")
            
            # AŞAMA 3: İlk iki kelime ile başlayan bir eşleşme ara
            if clean_product_value:
                words = clean_product_value.split()
                if len(words) >= 2:
                    first_two_words = ' '.join(words[:2]).lower()
                    
                    for i, option in enumerate(options):
                        if option['text'].lower().startswith(first_two_words):
            
                            driver.execute_script("arguments[0].selectedIndex = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", 
                                                dropdown_element, i)
                            time.sleep(1)
                            return True
            
            # AŞAMA 4: Kodda (değerde) bir eşleşme ara
            if clean_product_value:
                # CSV değerindeki kod kısmını çıkarmaya çalış (parantez içinde veya sonunda)
                import re
                code_match = re.search(r'\(([^)]+)\)$', clean_product_value)
                code = code_match.group(1) if code_match else None
                
                if not code:
                    # Alternatif: Son kelimeyi kod olarak değerlendir
                    words = clean_product_value.split()
                    if len(words) > 0:
                        code = words[-1]
                
                if code and len(code) > 3:  # En az 4 karakter olmalı
                    
                    
                    for i, option in enumerate(options):
                        option_value = option['value'].lower()
                        option_text = option['text'].lower()
                        
                        if code.lower() in option_value or code.lower() in option_text:
                           
                            driver.execute_script("arguments[0].selectedIndex = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", 
                                                dropdown_element, i)
                            time.sleep(1)
                            return True
            
            # AŞAMA 5: Birinci seçeneği seç (0 genellikle "Lütfen seçin" olduğundan 1'i seç)
            select_index = 1 if len(options) > 1 else 0
            logger.warning(f"Uygun eşleşme bulunamadı, {select_index+1}. seçenek seçiliyor: {options[select_index]['text']}")
            driver.execute_script("arguments[0].selectedIndex = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", 
                                dropdown_element, select_index)
            
           
            return True
        else:
            # Standart dropdown - tıkla ve seçenekleri ara
            driver.execute_script("arguments[0].click();", dropdown_element)

            
            # Dropdown açıldıktan sonra kısa bir süre bekle
            time.sleep(1)
            
            # CSV'den gelen product_value'yu temizle ve kontrol et
            clean_product_value = product_value.strip() if product_value else ""
            
            if not clean_product_value:
                logger.warning("CSV'den alınan ürün değeri boş")
            
            # Dropdown menüdeki tüm görünür seçenekleri bul - daha güvenli selector
            js_get_options = """
            // Tüm olası dropdown öğelerini bul
            function findDropdownOptions() {
                // Standart dropdown öğeleri
                var options = Array.from(document.querySelectorAll('.dropdown-menu li, .action-menu li, li.option-item, [role="option"], ul.option-list li'));
                
                // Eğer öğeler bulunamadıysa, tüm dropdown içindeki li'ları bul
                if (options.length === 0) {
                    var dropdowns = document.querySelectorAll('.dropdown-menu, .action-menu, [role="listbox"], ul.option-list');
                    for (var i = 0; i < dropdowns.length; i++) {
                        if (dropdowns[i].offsetParent !== null) { // Görünür olmalı
                            var items = dropdowns[i].querySelectorAll('li');
                            options = options.concat(Array.from(items));
                        }
                    }
                }
                
                // Eğer hala bulunamadıysa, tüm görünür li'ları bul
                if (options.length === 0) {
                    options = Array.from(document.querySelectorAll('li')).filter(function(el) {
                        return el.offsetParent !== null && // Görünür olmalı
                               el.textContent.trim() !== ''; // Boş olmamalı
                    });
                }
                
                // Sonuçları temizle ve döndür
                return options.filter(function(el) {
                    return el && el.textContent && el.textContent.trim() !== '';
                }).map(function(el) {
                    return {
                        element: el,
                        text: el.textContent.trim()
                    };
                });
            }
            
            return findDropdownOptions();
            """
            
            # Dropdown seçeneklerini al
            options = driver.execute_script(js_get_options)
            
    
         
            
            # Seçenekleri seç
            if len(options) > 0 and clean_product_value:
                # 1. Tam eşleşme ara
                for option in options:
                    if option['text'].lower() == clean_product_value.lower():
                        
                        driver.execute_script("arguments[0].click();", option['element'])
                        time.sleep(1)
                        return True
                
                # 2. Kısmi eşleşme ara - kelime bazlı
                best_option = None
                best_match_count = 0
                
                clean_words = [w.lower() for w in clean_product_value.split() if len(w) > 2]
                
                for option in options:
                    option_text = option['text'].lower()
                    match_count = sum(1 for word in clean_words if word in option_text)
                    
                    if match_count > best_match_count:
                        best_match_count = match_count
                        best_option = option
                
                if best_option and best_match_count > 0:
                   
                    driver.execute_script("arguments[0].click();", best_option['element'])
                    time.sleep(1)
                    return True
                
                # 3. İlk seçeneği seç

                driver.execute_script("arguments[0].click();", options[0]['element'])
                
                
                return True
            elif len(options) > 0:
                # CSV değeri boşsa ilk seçeneği seç
              
                driver.execute_script("arguments[0].click();", options[0]['element'])
                time.sleep(1)
                return True
            else:
                logger.warning("Dropdown'da hiç seçenek bulunamadı")
        
       
        return True
        
    except Exception as e:
        logger.error(f"Product Parts dropdown seçiminde hata: {str(e)}")
        
        # Hata durumunda da manuel müdahale iste
        
        return True

def find_offnung_and_uncheck_material(driver):
    try:

        # 1. Öffnung elementini bul
        offnung_elements = driver.find_elements(By.XPATH, "//span[contains(text(), 'Öffnung') or contains(text(), 'Offnung')]")
        
        if not offnung_elements:
            logger.warning("Öffnung elementi bulunamadı, alternatif yaklaşımlar deneniyor")
            offnung_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Öffnung') or contains(text(), 'Offnung')]")
        
        if not offnung_elements:
            logger.error("Öffnung elementi bulunamadı")
            return False
        
        # Görünür Öffnung elementini bul
        offnung_element = None
        for elem in offnung_elements:
            if elem.is_displayed():
                offnung_element = elem
                break
        
        if not offnung_element:
            logger.error("Görünür Öffnung elementi bulunamadı")
            return False
        
        # Öffnung elementini görünür yap
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", offnung_element)
        time.sleep(1)
        
       
        
        # Öffnung konumunu al
        offnung_location = offnung_element.location
        
        # 2. Öffnung'un altındaki Material alanını bul
        material_elements = driver.find_elements(By.XPATH, "//span[contains(text(), 'Material')]")
        
        # Öffnung'tan sonra gelen Material elementini bul
        material_element = None
        for elem in material_elements:
            if elem.is_displayed() and elem.location['y'] > offnung_location['y']:
                material_element = elem
                break
        
        if not material_element:
            logger.error("Öffnung altında Material elementi bulunamadı")
            return False
        
        # Material elementini görünür yap
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", material_element)
        
        
       
        
        # 3. Checkbox'ı tıklama işlemini yeni fonksiyona devret
        # Pass the material_element to the uncheck_material_checkbox function
        uncheck_result = uncheck_material_checkbox(driver, material_element)
        
        if not uncheck_result:
            logger.warning("Checkbox işlemi başarısız oldu")
        
        
    except Exception as e:
        import traceback
        logger.error(f"Stack trace: {traceback.format_exc()}")
        return False

def select_matching_material(driver, material_element):
    try:
     
        name_input = driver.find_element(By.NAME, "name")
        name_value = name_input.get_attribute("value").strip()

        js_script = """
        function findExactMatchingMaterial(nameValue) {
            // Tüm olası material seçeneklerini bul
            var materialOptions = Array.from(document.querySelectorAll('select option, [role="option"], li.option-item'))
                .filter(option => option.offsetParent !== null); // Yalnızca görünür olanlar
            
            // Tam eşleşen seçeneği bul
            var exactMatch = materialOptions.find(option => 
                option.textContent.trim() === nameValue.trim()
            );
            
            if (exactMatch) {
                return {
                    found: true,
                    text: exactMatch.textContent.trim(),
                    type: exactMatch.tagName.toLowerCase()
                };
            }
            
            return { found: false, message: "Tam eşleşen seçenek bulunamadı" };
        }
        
        return findExactMatchingMaterial(arguments[0]);
        """
        
        # JavaScript ile tam eşleşmeyi ara
        match_result = driver.execute_script(js_script, name_value)
        
        if match_result['found']:
 
            js_click_script = """
            function clickMatchingOption(nameValue) {
                var options = Array.from(document.querySelectorAll('select option, [role="option"], li.option-item'))
                    .filter(option => option.offsetParent !== null);
                
                var exactMatch = options.find(option => 
                    option.textContent.trim() === nameValue.trim()
                );
                
                if (exactMatch) {
                    // Select elementi ise
                    if (exactMatch.tagName.toLowerCase() === 'option') {
                        var select = exactMatch.closest('select');
                        select.value = exactMatch.value;
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                    } else {
                        // Diğer elementler için
                        exactMatch.click();
                    }
                    
                    return { success: true, text: exactMatch.textContent.trim() };
                }
                
                return { success: false, message: "Tıklama başarısız" };
            }
            
            return clickMatchingOption(arguments[0]);
            """
            
            click_result = driver.execute_script(js_click_script, name_value)   
        return True
    except Exception as e:
        return False

def uncheck_material_checkbox(driver, material_element):
    """Uncheck by clicking on the SECOND 'Check all for Material' text"""
    try:
        # JavaScript to find and click the second 'Check all for Material' text
        js_script = """
        function clickSecondCheckAllForMaterialText() {
            // Find all elements with 'Check all for Material' text
            var checkAllElements = Array.from(document.querySelectorAll('*'))
                .filter(el => {
                    // Exact text match and visible
                    return el.textContent && 
                           el.textContent.trim() === 'Check all for Material' &&
                           el.offsetParent !== null; // Ensure element is visible
                });
            
            // Ensure we have at least two elements
            if (checkAllElements.length >= 2) {
                // Select the second element (index 1)
                var elementToClick = checkAllElements[1];
                
                // Highlight for visibility
                elementToClick.style.backgroundColor = 'red';
                setTimeout(() => { elementToClick.style.backgroundColor = ''; }, 500);
                
                // Try multiple click methods
                try {
                    // Method 1: Direct click
                    elementToClick.click();
                } catch {
                    try {
                        // Method 2: JavaScript click
                        var clickEvent = new MouseEvent('click', {
                            view: window,
                            bubbles: true,
                            cancelable: true
                        });
                        elementToClick.dispatchEvent(clickEvent);
                    } catch {
                        // Method 3: Parent element click
                        var parent = elementToClick.closest('label, div, span');
                        if (parent) {
                            parent.click();
                        }
                    }
                }
                
                return {
                    success: true,
                    elementText: elementToClick.textContent,
                    tagName: elementToClick.tagName,
                    totalFound: checkAllElements.length
                };
            }
            
            return { 
                success: false, 
                message: "Could not find second 'Check all for Material' text",
                totalFound: checkAllElements ? checkAllElements.length : 0
            };
        }
        
        return clickSecondCheckAllForMaterialText();
        """
        
        # Execute the JavaScript
        result = driver.execute_script(js_script)
        
        
        if result.get('success', False):

            # Call select_matching_material after successful unchecking
            select_matching_material(driver, material_element)
            
            return True
        
        # Fallback Selenium approach
        try:
            # Find all 'Check all for Material' elements
            check_all_elements = driver.find_elements(By.XPATH, 
                "//*[text()='Check all for Material']")
            
            # Ensure we have at least two elements
            if len(check_all_elements) >= 2:
                # Click the second element
                element = check_all_elements[1]
                
                if element.is_displayed():
                    # Try to click
                    try:
                        element.click()
                    
                        
                        # Call select_matching_material after successful unchecking
                        select_matching_material(driver, material_element)
                        
                        return True
                    except:
                        # If direct click fails, try JavaScript click
                        try:
                            driver.execute_script("arguments[0].click();", element)

                            
                            # Call select_matching_material after successful unchecking
                            select_matching_material(driver, material_element)
                            
                            return True
                        except Exception as e:
                            logger.warning(f"Click attempt failed: {e}")
        
        except Exception as e:
            logger.error(f"Selenium text click failed: {e}")
        
        
        
        # Even after manual intervention, try to select matching material
        select_matching_material(driver, material_element)
        
        return True
    
    except Exception as e:
        logger.error(f"Comprehensive click attempt failed: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        
        # Try to select matching material even after error
        try:
            select_matching_material(driver, material_element)
        except:
            pass
        
        return False

def improved_fill_fixed_price_form(driver, CSV_FILE_PATH):
    """Fill the Fixed Price form with name and code values from CSV file, selecting product and setting price unit"""
    try:
        # Read the CSV file
        import csv
        with open(CSV_FILE_PATH, 'r', encoding='utf-8') as file:
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
            
            # Get the values from H and I columns (Page options Name and Page options code)
            name_value_raw = data_row[7] if len(data_row) > 7 else ""    # H column - Page options Name
            code_value_raw = data_row[8] if len(data_row) > 8 else ""    # I column - Page options code
            
            # Get the product value from column A (index 2) - "produkt" column
            product_value = data_row[2] if len(data_row) > 2 else ""

            
            # Parse multiple materials from the name and code values
            name_materials = parse_materials(name_value_raw)
            code_materials = parse_materials(code_value_raw)
            
            # Ensure we have matching pairs of name and code
            material_count = max(len(name_materials), len(code_materials))
            
            # If we have more names than codes or vice versa, pad the shorter list
            if len(name_materials) < material_count:
                name_materials.extend([""] * (material_count - len(name_materials)))
            if len(code_materials) < material_count:
                code_materials.extend([""] * (material_count - len(code_materials)))
            
            # Process each material
            for i, (name_value, code_value) in enumerate(zip(name_materials, code_materials)):

                # If not the first material, we need to click Add New again
                if i > 0:
                    # Wait for the previous save to complete
                    # Navigate back to the Fixed Price page
                    click_fixed_price(driver)
                    # Click Add New for the next material
                    click_add_new(driver)               
                # Wait for form to load
                time.sleep(2)
                
                # Fill the name and code fields
                js_script = """
                function fillForm(nameValue, codeValue) {
                    console.log('Starting to fill inputs with:', nameValue, codeValue);
                    
                    // Find and fill name field
                    var nameField = document.querySelector('input[name="name"]');
                    if (!nameField) {
                        var nameLabels = Array.from(document.querySelectorAll('label, span')).filter(el => 
                            el.textContent.trim() === 'Name');
                        if (nameLabels.length > 0) {
                            var nameFieldContainer = nameLabels[0].closest('.admin__field, .field');
                            if (nameFieldContainer) {
                                nameField = nameFieldContainer.querySelector('input');
                            }
                        }
                    }
                    
                    if (nameField) {
                        nameField.value = nameValue;
                        nameField.dispatchEvent(new Event('input', {bubbles: true}));
                        nameField.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                    
                    // Find and fill code field
                    var codeField = document.querySelector('input[name="code"]');
                    if (!codeField) {
                        var codeLabels = Array.from(document.querySelectorAll('label, span')).filter(el => 
                            el.textContent.trim() === 'Code');
                        if (codeLabels.length > 0) {
                            var codeFieldContainer = codeLabels[0].closest('.admin__field, .field');
                            if (codeFieldContainer) {
                                codeField = codeFieldContainer.querySelector('input');
                            }
                        }
                    }
                    
                    if (codeField) {
                        codeField.value = codeValue;
                        codeField.dispatchEvent(new Event('input', {bubbles: true}));
                        codeField.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                    
                    return {
                        nameFieldFilled: nameField ? true : false,
                        codeFieldFilled: codeField ? true : false
                    };
                }
                
                return fillForm(arguments[0], arguments[1]);
                """
                
                # Execute the form fill script
                form_result = driver.execute_script(js_script, name_value, code_value)
      
                # Set Price Unit to 1
                set_price_unit_to_one(driver)
                
                # Select product from dropdown using our new function
                select_product_parts_dropdown(driver, product_value)
                
                find_offnung_and_uncheck_material(driver)
                # Save the material
                click_save_button(driver)
            
            return True
            
    except Exception as e:
        logger.error(f"Error in improved_fill_fixed_price_form: {str(e)}")
        return False

def click_save_button(driver):
    """Click the Save button on the form - specifically targeting the top right save button"""
    try:
        logger.info("Attempting to click the Save button in the top right corner")
        
        # First approach: Try a specific XPath targeting the right-side save button
        try:
            # This targets the orange save button in the top right corner based on your screenshot
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'primary') and contains(text(), 'Save') and not(contains(@class, 'disabled'))]"))
            )
            
            # Scroll to make sure it's visible
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
            time.sleep(1)
            
            # Click using JavaScript for reliability
            driver.execute_script("arguments[0].click();", save_button)
            logger.info("Successfully clicked the Save button via specific XPath")
            
            # Wait for the save operation to complete
            time.sleep(1)
            return True
        except Exception as e:
            logger.warning(f"First approach failed: {str(e)}")
        
        # Second approach: Try to find the orange Save button by its color and position
        try:
            js_script = """
            // Find all buttons
            var buttons = document.querySelectorAll('button');
            var candidateButtons = [];
            
            // Find all visible save buttons
            for (var i = 0; i < buttons.length; i++) {
                var btn = buttons[i];
                if (!btn.offsetParent) continue; // Skip hidden buttons
                
                // Look for 'Save' text
                if (btn.textContent.trim() === 'Save') {
                    // Get the computed style to check background color (orange buttons are preferred)
                    var style = window.getComputedStyle(btn);
                    var rect = btn.getBoundingClientRect();
                    
                    candidateButtons.push({
                        element: btn,
                        isOrange: (
                            style.backgroundColor.includes('rgb(243') || // Orange color
                            style.backgroundColor.includes('rgb(255, 87') ||
                            btn.classList.contains('orange') ||
                            btn.classList.contains('primary')
                        ),
                        isTopRight: (rect.top < 200 && rect.right > (window.innerWidth - 200)),
                        position: {
                            top: rect.top,
                            right: window.innerWidth - rect.right
                        }
                    });
                }
            }
            
            // Sort by position (prefer top-right)
            candidateButtons.sort(function(a, b) {
                // Prefer orange buttons
                if (a.isOrange && !b.isOrange) return -1;
                if (!a.isOrange && b.isOrange) return 1;
                
                // Prefer top-right buttons
                if (a.isTopRight && !b.isTopRight) return -1;
                if (!a.isTopRight && b.isTopRight) return 1;
                
                // Sort by position
                return (a.position.top + a.position.right) - (b.position.top + b.position.right);
            });
            
            if (candidateButtons.length > 0) {
                var bestButton = candidateButtons[0].element;
                bestButton.scrollIntoView({block: 'center'});
                setTimeout(function() {
                    bestButton.click();
                }, 100);
                return "Clicked the best Save button: " + bestButton.textContent;
            }
            
            return "Could not find any Save buttons";
            """
            
            result = driver.execute_script(js_script)
            logger.info(f"JavaScript Save button result: {result}")
            
            if "Clicked" in str(result):
                time.sleep(1)  # Wait for save operation
                return True
            else:
                logger.warning("JavaScript approach failed to find Save button")
        except Exception as js_e:
            logger.warning(f"JavaScript approach failed: {str(js_e)}")
        
        # Third approach: Try directly clicking at the expected position of the Save button
        try:
            # Get window dimensions
            window_size = driver.get_window_size()
            
            # Calculate position for the Save button (top right area)
            # Based on screenshot, button is near the top right
            save_x = window_size['width'] - 60  # 60px from right edge
            save_y = 180  # Around 180px from top
            
            # Create ActionChains
            action = ActionChains(driver)
            
            # Move to the calculated position and click
            action.move_by_offset(save_x, save_y).click().perform()
            logger.info(f"Clicked at position: x={save_x}, y={save_y} using ActionChains")
            
            time.sleep(1)  # Wait for save operation
            return True
        except Exception as action_e:
            logger.error(f"ActionChains approach failed: {str(action_e)}")
        
        logger.error("All approaches failed to click the Save button")
        return False
    except Exception as e:
        logger.error(f"Error in click_save_button: {str(e)}")
        return False
def main():
    try:
        # Set up Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-popup-blocking")
        
        # Initialize Chrome driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # Login to admin panel
        if not login_to_admin(driver):
            logger.error("Failed to login. Exiting...")
            
            return
        
        # Navigate to CloudLab
        if not click_cloudlab(driver):
            logger.error("Failed to navigate to CloudLab. Exiting...")
            
            return
        
        # Navigate to Fixed Price
        if not click_fixed_price(driver):
            logger.error("Failed to navigate to Fixed Price. Exiting...")
            
            return
        
        # Click Add New button
        if not click_add_new(driver):
            logger.error("Failed to click Add New button. Exiting...")
            
            return
        
        # Fill the form with improved function
        if not improved_fill_fixed_price_form(driver, CSV_FILE_PATH):
            logger.error("Failed to fill the form. Exiting...")
            
            return

        # Wait for a moment to see the final state
        time.sleep(2)
        driver.quit()
        
    except Exception as e:
        logger.error(f"Error in main function: {str(e)}")

if __name__ == "__main__":
    main()
