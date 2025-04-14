"""
PrintNet Automation Module
"""

import csv
import time
import os
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# Setup logging
def setup_logging():
    """Set up logging configuration"""
    # Create logs directory if it doesn't exist
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    # Set up log file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"printnet_automation_{timestamp}.log")
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()  # Also output to console
        ]
    )
    
    logging.info(f"Logging initialized. Log file: {log_file}")
    return logging.getLogger()

# Initialize logger
logger = setup_logging()
username = "Serhat.Seyitoglu" 
Password = "Serhat.kamile.mavi2212" 

# Configuration
CSV_FILE_PATH = "C:/Users/g.yilmaz/Desktop/Magento_otomation/csv_files/test.csv"
PRINTNET_URL = "https://printnet.networkhero.eu/dashboard/admin/"  # PrintNet URL'inizi güncelleyin

CSV_MAPPINGS = {
    "product_name": 2,  # Column C (index 2 in CSV)
    "code": 4,          # Column E (index 4 in CSV)
    "sku": 6,        # Column G (index 6 in CSV)
    "product_part_options": 10,  # Bu sütun numarasını CSV dosyanızdaki gerçek konuma göre ayarlayın
    "product_options": 11        # Bu sütun numarasını CSV dosyanızdaki gerçek konuma göre ayarlayın
}

# Fixed values
PRICE = "1.00"
QUANTITY = "99999999"
VPE_VALUE = "9999"
VPE_TYPE = "Stock"  # New setting for VPE Type
COMBINATION_OPTION = "Fix Products"

def login_to_printnet(driver, username, password):
    """Login to PrintNet system"""
    try:
        logger.info(f"Navigating to PrintNet login page: {PRINTNET_URL}")
        driver.get(PRINTNET_URL)
        
        
        # Enter username - ID'yi PrintNet'in gerçek login sayfası yapısına göre güncelleyin
        try:
            username_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            username_field.clear()
            username_field.send_keys(username)
            logger.info("Username entered")
        except Exception as e:
            logger.error(f"Failed to find username field: {str(e)}")
            
            # Try alternative selectors for username field
            try:
                username_field = driver.find_element(By.CSS_SELECTOR, "input[name='username'], input[type='text']")
                username_field.clear()
                username_field.send_keys(username)
                logger.info("Username entered using alternative selector")
            except Exception as alt_e:
                logger.error(f"All attempts to find username field failed: {str(alt_e)}")
                return False
        
        # Enter password - ID'yi PrintNet'in gerçek login sayfası yapısına göre güncelleyin
        try:
            password_field = driver.find_element(By.ID, "password")
            password_field.clear()
            password_field.send_keys(password)
            logger.info("Password entered")
        except Exception as e:
            logger.error(f"Failed to find password field by ID: {str(e)}")
            
            # Try alternative selectors for password field
            try:
                password_field = driver.find_element(By.CSS_SELECTOR, "input[name='password'], input[type='password']")
                password_field.clear()
                password_field.send_keys(password)
                logger.info("Password entered using alternative selector")
            except Exception as alt_e:
                logger.error(f"All attempts to find password field failed: {str(alt_e)}")
                return False
        
        
        # Click login button - PrintNet'in gerçek login butonu seçicisini kullanın
        try:
            login_button = driver.find_element(By.CSS_SELECTOR, ".action-login, .login-button, button[type='submit']")
            driver.execute_script("arguments[0].click();", login_button)
            logger.info("Login button clicked")
        except Exception as e:
            logger.error(f"Failed to find login button: {str(e)}")
            
            # Try by finding any button on the login form
            try:
                login_buttons = driver.find_elements(By.CSS_SELECTOR, "button")
                if login_buttons:
                    driver.execute_script("arguments[0].click();", login_buttons[0])
                    logger.info("Clicked first button found on login form")
                else:
                    logger.error("No buttons found on login form")
                    return False
            except Exception as alt_e:
                logger.error(f"All attempts to find login button failed: {str(alt_e)}")
                return False
        
        # Wait for login to complete - PrintNet'in ana sayfa göstergesini güncelleyin
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".page-header, .dashboard, .admin-user"))
            )
            logger.info("Successfully logged in to PrintNet")
            
          
            
            return True
        except Exception as e:
            logger.error(f"Failed to detect successful login: {str(e)}")
            
            
            
            return False
    except Exception as e:
        logger.error(f"Login to PrintNet failed: {str(e)}")
        return False

def navigate_to_add_product_printnet(driver):
    """Navigate to the Add Product page in PrintNet"""
    try:
        # Click on Catalog icon in the menu
        logger.info("Clicking on Catalog icon in PrintNet...")
        catalog_icon = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-catalog')]//a"))
        )
        driver.execute_script("arguments[0].click();", catalog_icon)
        
        # Click on Products in the submenu
        logger.info("Clicking on Products submenu in PrintNet...")
        products_submenu = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'item-catalog')]//li[contains(@class, 'item-catalog-products')]//a"))
        )
        driver.execute_script("arguments[0].click();", products_submenu)
        
        # Wait for products page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
        )
        
        # Additional wait to ensure page is fully loaded
        # Click on Add Product button
        logger.info("Clicking on Add Product button in PrintNet...")
        add_product_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.add"))
        )
        driver.execute_script("arguments[0].click();", add_product_button)
        
        # Wait for product creation page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".page-title"))
        )
        
        logger.info("Navigated to Add Product page in PrintNet")
        return True
    except Exception as e:
        logger.error(f"Error navigating to Add Product page in PrintNet: {str(e)}")
        return False

def fill_basic_info_printnet(driver, product_name, sku):
    """CSV'den alınan verilerle temel ürün bilgilerini doldur"""
    try:
        
        
        # Enter product name from CSV
        logger.info(f"Ürün adı giriliyor: {product_name}")
        try:
            product_name_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='product[name]']"))
            )
            product_name_field.clear()
            product_name_field.send_keys(product_name)
            logger.info(f"Ürün adı girildi: {product_name}")
        except Exception as name_e:
            logger.error(f"Ürün adı alanı bulunamadı: {str(name_e)}")
            
            # JavaScript alternatifi
            try:
                result = driver.execute_script("""
                    // Ürün adı alanını bul
                    var nameFields = Array.from(document.querySelectorAll('input')).filter(input => 
                        input.name && (
                            input.name === 'product[name]' ||
                            input.id === 'product_name' ||
                            input.placeholder === 'Product Name' || 
                            input.placeholder === 'Name'
                        )
                    );
                    
                    if (nameFields.length > 0) {
                        nameFields[0].value = arguments[0];
                        nameFields[0].dispatchEvent(new Event('input', {bubbles: true}));
                        nameFields[0].dispatchEvent(new Event('change', {bubbles: true}));
                        return "Ürün adı JavaScript ile girildi";
                    }
                    
                    return "Ürün adı alanı bulunamadı";
                """, product_name)
                
                logger.info(f"JavaScript sonucu: {result}")
                if "bulunamadı" in result:
                    return False
            except Exception as js_e:
                logger.error(f"JavaScript ile ürün adı girilirken hata: {str(js_e)}")
                return False
        
        # Enter SKU from CSV
        logger.info(f"SKU giriliyor: {sku}")
        try:
            sku_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='product[sku]']"))
            )
            sku_field.clear()
            sku_field.send_keys(sku)
            logger.info(f"SKU girildi: {sku}")
        except Exception as sku_e:
            logger.error(f"SKU alanı bulunamadı: {str(sku_e)}")
            
            # JavaScript alternatifi
            try:
                result = driver.execute_script("""
                    // SKU alanını bul
                    var skuFields = Array.from(document.querySelectorAll('input')).filter(input => 
                        input.name && (
                            input.name === 'product[sku]' ||
                            input.id === 'product_sku' ||
                            input.placeholder === 'SKU'
                        )
                    );
                    
                    if (skuFields.length > 0) {
                        skuFields[0].value = arguments[0];
                        skuFields[0].dispatchEvent(new Event('input', {bubbles: true}));
                        skuFields[0].dispatchEvent(new Event('change', {bubbles: true}));
                        return "SKU JavaScript ile girildi";
                    }
                    
                    return "SKU alanı bulunamadı";
                """, sku)
                
                logger.info(f"JavaScript sonucu: {result}")
                if "bulunamadı" in result:
                    return False
            except Exception as js_e:
                logger.error(f"JavaScript ile SKU girilirken hata: {str(js_e)}")
                return False
        
        # Enter price (sabit değer)
        logger.info(f"Fiyat giriliyor: {PRICE}")
        try:
            price_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='product[price]']"))
            )
            price_field.clear()
            price_field.send_keys(PRICE)
            logger.info(f"Fiyat girildi: {PRICE}")
        except Exception as price_e:
            logger.error(f"Fiyat alanı bulunamadı: {str(price_e)}")
            
            # JavaScript alternatifi
            try:
                result = driver.execute_script("""
                    // Fiyat alanını bul
                    var priceFields = Array.from(document.querySelectorAll('input')).filter(input => 
                        input.name && (
                            input.name === 'product[price]' ||
                            input.id === 'product_price' ||
                            input.placeholder === 'Price'
                        )
                    );
                    
                    if (priceFields.length > 0) {
                        priceFields[0].value = arguments[0];
                        priceFields[0].dispatchEvent(new Event('input', {bubbles: true}));
                        priceFields[0].dispatchEvent(new Event('change', {bubbles: true}));
                        return "Fiyat JavaScript ile girildi";
                    }
                    
                    return "Fiyat alanı bulunamadı";
                """, PRICE)
                
                logger.info(f"JavaScript sonucu: {result}")
                if "bulunamadı" in result:
                    return False
            except Exception as js_e:
                logger.error(f"JavaScript ile fiyat girilirken hata: {str(js_e)}")
                return False
        
        # Enter quantity (sabit değer)
        logger.info(f"Miktar giriliyor: {QUANTITY}")
        try:
            qty_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='product[quantity_and_stock_status][qty]']"))
            )
            qty_field.clear()
            qty_field.send_keys(QUANTITY)
            logger.info(f"Miktar girildi: {QUANTITY}")
        except Exception as qty_e:
            logger.error(f"Miktar alanı bulunamadı: {str(qty_e)}")
            
            # JavaScript alternatifi
            try:
                result = driver.execute_script("""
                    // Miktar alanını bul
                    var qtyFields = Array.from(document.querySelectorAll('input')).filter(input => 
                        input.name && (
                            input.name === 'product[quantity_and_stock_status][qty]' ||
                            input.id === 'product_qty' ||
                            input.placeholder === 'Quantity'
                        )
                    );
                    
                    if (qtyFields.length > 0) {
                        qtyFields[0].value = arguments[0];
                        qtyFields[0].dispatchEvent(new Event('input', {bubbles: true}));
                        qtyFields[0].dispatchEvent(new Event('change', {bubbles: true}));
                        return "Miktar JavaScript ile girildi";
                    }
                    
                    return "Miktar alanı bulunamadı";
                """, QUANTITY)
                
                logger.info(f"JavaScript sonucu: {result}")
                if "bulunamadı" in result:
                    return False
            except Exception as js_e:
                logger.error(f"JavaScript ile miktar girilirken hata: {str(js_e)}")
                return False
        
    
        
        logger.info(f"Temel bilgiler başarıyla dolduruldu: {product_name} (SKU: {sku})")
        return True
    except Exception as e:
        logger.error(f"Temel bilgiler doldurulurken hata: {str(e)}")
        return False

def set_visibility_to_catalog(driver):
    """
    Sets the Visibility option to 'Catalog'
    
    Args:
        driver: Selenium WebDriver instance
    
    Returns:
        bool: True if successfully set to 'Catalog', False otherwise
    """
    try:
        logger.info("Starting Visibility selection process")
        
        # Find the Visibility dropdown directly
        visibility_dropdown_selectors = [
            "//select[@name='product[visibility]']",  # By name attribute
            "//div[contains(@class, 'admin__field')]//select[contains(@data-ui-id, 'visibility')]",  # By data attribute
            "//div[contains(text(), 'Visibility')]/following-sibling::div//select"  # By text proximity
        ]
        
        dropdown = None
        for selector in visibility_dropdown_selectors:
            try:
                dropdown = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                if dropdown:
                    break
            except:
                continue
        
        if not dropdown:
            logger.error("Could not find Visibility dropdown")
            return False
        
        # Create Select object to interact with dropdown
        from selenium.webdriver.support.ui import Select
        select = Select(dropdown)
        
        # Select 'Catalog' option
        try:
            # Try to select by visible text
            select.select_by_visible_text('Catalog')
            logger.info("Successfully selected 'Catalog'")
            time.sleep(1)
            return True
        
        except Exception as e:
            logger.error(f"Failed to select 'Catalog': {str(e)}")
            
            # Fallback: try selecting by value or index if text selection fails
            try:
                # Try selecting by value if visible text fails
                select.select_by_value('catalog')
                logger.info("Selected 'Catalog' by value")
                time.sleep(1)
                return True
            except:
                # Last resort: try selecting by index if other methods fail
                try:
                    # Typically, 'Catalog' might be the second or third option
                    select.select_by_index(1)  # Try second option
                    logger.info("Selected 'Catalog' by index")
                    time.sleep(1)
                    return True
                except Exception as index_e:
                    logger.error(f"Failed to select 'Catalog': {str(index_e)}")
                    return False
    
    except Exception as main_e:
        logger.error(f"Overall failure in setting Visibility to Catalog: {str(main_e)}")
        return False

def set_meter_price_printnet(driver, is_meter_price=True):
    """Set the 'Is Meter Price' option according to the parameter value with advanced positioning"""
    try:
        logger.info(f"Setting 'Is Meter Price' to: {'Yes' if is_meter_price else 'No'}")
        
     
        # First, disable smooth scrolling to avoid animation issues
        driver.execute_script("document.documentElement.style.scrollBehavior = 'auto';")
        
        # Store initial scroll position to check if page jumps
        initial_scroll_y = driver.execute_script("return window.pageYOffset;")
        logger.info(f"Initial scroll position: {initial_scroll_y}px")
        
        # Execute a direct search for the Is Meter Price element using JavaScript
        # This avoids scrolling and element focus issues
        meter_price_position = driver.execute_script("""
            // Find all elements with text "Is Meter Price"
            var elements = Array.from(document.querySelectorAll('*')).filter(el => 
                el.textContent && el.textContent.trim() === 'Is Meter Price');
            
            if (elements.length === 0) return null;
            
            // Get the one that's most likely the label (span within admin field)
            var label = elements.find(el => el.tagName.toLowerCase() === 'span');
            if (!label) label = elements[0];
            
            // Return the element's position
            var rect = label.getBoundingClientRect();
            return {
                x: rect.left + window.pageXOffset,
                y: rect.top + window.pageYOffset,
                found: true
            };
        """)
        
        if not meter_price_position or not meter_price_position.get('found'):
            logger.error("Could not find 'Is Meter Price' label with JavaScript")
            raise Exception("Is Meter Price label not found")
        
        # Scroll to position with a slight offset to ensure it's visible
        # Use absolute positioning to avoid relative scrolling issues
        y_position = meter_price_position['y'] - 150  # Added offset to ensure good visibility
        driver.execute_script(f"window.scrollTo(0, {y_position});")
        
        
        # Verify we're at expected position
        current_scroll_y = driver.execute_script("return window.pageYOffset;")
        logger.info(f"Current scroll position after positioning: {current_scroll_y}px")
        
     
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

def process_printnet(driver):
    """
    Ürün sayfasındaki website seçimlerini yapar:
    1. Sayfayı aşağı kaydırarak "Product in Websites" yazısını bulur ve tıklar
    2. bannerheld.de bulup tıklar (işareti kaldırır)
    3. bannerheld.at bulup tıklar (işaretler)
    """
    try:
        logger.info("Website seçimlerini yapıyorum...")
        
        # 1. "Product in Websites" başlığını bul ve tıkla
        # Sayfayı adım adım aşağı kaydırarak elementleri ara
        product_in_websites_found = False
        scroll_attempts = 0
        max_scroll_attempts = 20  # Maksimum kaydırma denemesi
        
        while not product_in_websites_found and scroll_attempts < max_scroll_attempts:
            # Şu anki sayfada "Product in Websites" başlığını aramaya çalış
            try:
                # Görünür elementleri arama
                websites_elements = driver.find_elements(By.XPATH, "//div[contains(text(), 'Product in Websites')] | //span[contains(text(), 'Product in Websites')]")
                
                # Görünür bir element bulundu mu kontrol et
                for element in websites_elements:
                    if element.is_displayed():
                        # Element görünür, kaydır ve tıkla
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                    
                        
                        
                        # Elemente tıkla
                        driver.execute_script("arguments[0].click();", element)
                        logger.info("Product in Websites başlığına tıklandı")
                        
                        product_in_websites_found = True
                        break
                
                if not product_in_websites_found:
                    # Element bulunamadı, sayfayı biraz daha aşağı kaydır
                    driver.execute_script("window.scrollBy(0, 300);")  # 300px aşağı kaydır
                    scroll_attempts += 1
                    
                    logger.info(f"Sayfayı aşağı kaydırıyorum - deneme {scroll_attempts}/{max_scroll_attempts}")
            
            except Exception as find_e:
                logger.warning(f"Element arama hatası: {str(find_e)}")
                # Sayfayı biraz daha aşağı kaydır ve tekrar dene
                driver.execute_script("window.scrollBy(0, 300);")
                scroll_attempts += 1
               
        
        # Eğer "Product in Websites" bulunamadıysa
        if not product_in_websites_found:
            logger.error("Product in Websites başlığı bulunamadı, son bir JavaScript denemesi yapılıyor...")
            
            # Son bir JavaScript denemesi
            result = driver.execute_script("""
                // Sayfadaki tüm elementleri kontrol et
                var elements = document.querySelectorAll('*');
                
                for (var i = 0; i < elements.length; i++) {
                    // "Product in Websites" içeren bir element ara
                    if (elements[i].textContent && elements[i].textContent.includes('Product in Websites')) {
                        // Elementi görünür yap
                        elements[i].scrollIntoView({block: 'center'});
                        
                        // 500ms bekle
                        setTimeout(function() {}, 500);
                        
                        // Elementi tıkla
                        elements[i].click();
                        return "Product in Websites element bulundu ve tıklandı";
                    }
                }
                
                return "Product in Websites elementi bulunamadı";
            """)
            
            logger.info(f"JavaScript sonucu: {result}")
            product_in_websites_found = "bulunamadı" not in result
            
           
        
        # Eğer hala bulunamadıysa, işlemi iptal et
        if not product_in_websites_found:
            logger.error("Product in Websites bölümü bulunamadı, işleme devam edilemiyor.")
            return False

        # 2. bannerheld.de checkbox'ını bul ve tıkla (işareti kaldır)
        # Önce görünür olması için kaydır
        bannerheld_de_found = False
        scroll_attempts = 0
        
        while not bannerheld_de_found and scroll_attempts < max_scroll_attempts:
            try:
                # bannerheld.de etiketini veya checkbox'ını ara
                de_elements = driver.find_elements(By.XPATH, "//label[contains(text(), 'Bannerheld.de')] | //input[@type='checkbox']/following-sibling::label[contains(text(), 'Bannerheld.de')]")
                
                # Görünür bir element bulundu mu kontrol et
                for element in de_elements:
                    if element.is_displayed():
                        # Elementi görünür hale getir
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                        
                        # Checkbox'ı bul
                        try:
                            # Önce kardeş element olarak dene
                            checkbox = element.find_element(By.XPATH, "./preceding-sibling::input[@type='checkbox']")
                        except:
                            # Yakındaki herhangi bir checkbox'ı bul
                            try:
                                checkbox = driver.find_element(By.XPATH, "//label[contains(text(), 'Bannerheld.de')]/preceding-sibling::input[@type='checkbox']")
                            except:
                                logger.warning("bannerheld.de checkbox'ı bulunamadı, JavaScript kullanılacak")
                                
                                # JavaScript ile bul ve tıkla
                                result = driver.execute_script("""
                                    var labels = document.querySelectorAll('label');
                                    
                                    for (var i = 0; i < labels.length; i++) {
                                        if (labels[i].textContent.includes('Bannerheld.de')) {
                                            var checkbox = labels[i].previousElementSibling;
                                            
                                            if (checkbox && checkbox.type === 'checkbox') {
                                                // İşaretli ise tıkla
                                                if (checkbox.checked) {
                                                    checkbox.click();
                                                    return "bannerheld.de checkbox tıklandı";
                                                } else {
                                                    return "bannerheld.de zaten işaretli değil";
                                                }
                                            }
                                        }
                                    }
                                    
                                    return "bannerheld.de checkbox bulunamadı";
                                """)
                                
                                logger.info(f"JavaScript sonucu: {result}")
                                bannerheld_de_found = "bulunamadı" not in result
                                break
                        
                        # Checkbox bulundu, durumunu kontrol et
                        if checkbox.is_selected():
                            # İşaretli ise, tıkla ve seçimi kaldır
                            driver.execute_script("arguments[0].click();", checkbox)
                            logger.info("bannerheld.de seçimi kaldırıldı")
                            
                        else:
                            logger.info("bannerheld.de zaten seçili değil")
                        
                        bannerheld_de_found = True
                        break
                
                if not bannerheld_de_found:
                    # Element bulunamadı, sayfayı biraz daha aşağı kaydır
                    driver.execute_script("window.scrollBy(0, 200);")
                    scroll_attempts += 1
 
            
            except Exception as find_e:
                logger.warning(f"bannerheld.de arama hatası: {str(find_e)}")
                # Sayfayı biraz daha aşağı kaydır ve tekrar dene
                driver.execute_script("window.scrollBy(0, 200);")
                scroll_attempts += 1

        
        # bannerheld.de bulunamadıysa JavaScript'i son çare olarak dene
        if not bannerheld_de_found:
            logger.warning("bannerheld.de normal yollarla bulunamadı, JavaScript denenecek")
            
            result = driver.execute_script("""
                
                
                // Tüm elementleri kontrol et ve bannerheld.de'yi bul
                var elements = document.querySelectorAll('*');
                
                for (var i = 0; i < elements.length; i++) {
                    if (elements[i].textContent && elements[i].textContent.includes('Bannerheld.de')) {
                        // Elementi görünür yap
                        elements[i].scrollIntoView({block: 'center'});
                        
                        // Checkbox'ı bul
                        var container = elements[i].closest('div');
                        var checkbox = container.querySelector('input[type="checkbox"]');
                        
                        if (checkbox) {
                            // Checkbox'ı işaretli ise kaldır
                            if (checkbox.checked) {
                                checkbox.click();
                                return "bannerheld.de işareti kaldırıldı";
                            }
                            return "bannerheld.de zaten işaretli değil";
                        }
                    }
                }
                
                return "bannerheld.de elementi bulunamadı";
            """)
            
            logger.info(f"JavaScript sonucu: {result}")
        
        # 3. bannerheld.at checkbox'ını bul ve tıkla (işaretle)
        # Önce görünür olması için kaydır
        bannerheld_at_found = False
        scroll_attempts = 0
        
        
        while not bannerheld_at_found and scroll_attempts < max_scroll_attempts:
            try:
                # bannerheld.at etiketini veya checkbox'ını ara
                at_elements = driver.find_elements(By.XPATH, "//label[contains(text(), 'Bannerheld.at')] | //input[@type='checkbox']/following-sibling::label[contains(text(), 'Bannerheld.at')]")
                
                # Görünür bir element bulundu mu kontrol et
                for element in at_elements:
                    if element.is_displayed():
                        # Elementi görünür hale getir
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                        
                       
                        
                        # Checkbox'ı bul
                        try:
                            # Önce kardeş element olarak dene
                            checkbox = element.find_element(By.XPATH, "./preceding-sibling::input[@type='checkbox']")
                        except:
                            # Yakındaki herhangi bir checkbox'ı bul
                            try:
                                checkbox = driver.find_element(By.XPATH, "//label[contains(text(), 'Bannerheld.at')]/preceding-sibling::input[@type='checkbox']")
                            except:
                                logger.warning("bannerheld.at checkbox'ı bulunamadı, JavaScript kullanılacak")
                                
                                # JavaScript ile bul ve tıkla
                                result = driver.execute_script("""
                                    var labels = document.querySelectorAll('label');
                                    
                                    for (var i = 0; i < labels.length; i++) {
                                        if (labels[i].textContent.includes('Bannerheld.at')) {
                                            var checkbox = labels[i].previousElementSibling;
                                            
                                            if (checkbox && checkbox.type === 'checkbox') {
                                                // İşaretli değilse tıkla
                                                if (!checkbox.checked) {
                                                    checkbox.click();
                                                    return "bannerheld.at checkbox tıklandı";
                                                } else {
                                                    return "bannerheld.at zaten işaretli";
                                                }
                                            }
                                        }
                                    }
                                    
                                    return "bannerheld.at checkbox bulunamadı";
                                """)
                                
                                logger.info(f"JavaScript sonucu: {result}")
                                bannerheld_at_found = "bulunamadı" not in result
                                break
                        
                        # Checkbox bulundu, durumunu kontrol et
                        if not checkbox.is_selected():
                            # İşaretli değilse, tıkla ve seç
                            driver.execute_script("arguments[0].click();", checkbox)
                            logger.info("bannerheld.at seçildi")
                            
                        else:
                            logger.info("bannerheld.at zaten seçili")
                        
                        bannerheld_at_found = True
                        break
                
                if not bannerheld_at_found:
                    # Element bulunamadı, sayfayı biraz daha aşağı kaydır
                    driver.execute_script("window.scrollBy(0, 200);")
                    scroll_attempts += 1

            except Exception as find_e:
                logger.warning(f"bannerheld.at arama hatası: {str(find_e)}")
                # Sayfayı biraz daha aşağı kaydır ve tekrar dene
                driver.execute_script("window.scrollBy(0, 200);")
                scroll_attempts += 1
                
        
        # bannerheld.at bulunamadıysa JavaScript'i son çare olarak dene
        if not bannerheld_at_found:
            logger.warning("bannerheld.at normal yollarla bulunamadı, JavaScript denenecek")
            
            result = driver.execute_script("""
                
                // Tüm elementleri kontrol et ve bannerheld.at'yi bul
                var elements = document.querySelectorAll('*');
                
                for (var i = 0; i < elements.length; i++) {
                    if (elements[i].textContent && elements[i].textContent.includes('Bannerheld.at')) {
                        // Elementi görünür yap
                        elements[i].scrollIntoView({block: 'center'});
                        
                        // Checkbox'ı bul
                        var container = elements[i].closest('div');
                        var checkbox = container.querySelector('input[type="checkbox"]');
                        
                        if (checkbox) {
                            // Checkbox'ı işaretli değilse işaretle
                            if (!checkbox.checked) {
                                checkbox.click();
                                return "bannerheld.at işaretlendi";
                            }
                            return "bannerheld.at zaten işaretli";
                        }
                    }
                }
                
                return "bannerheld.at elementi bulunamadı";
            """)
            
            logger.info(f"JavaScript sonucu: {result}")
        

        
        # Sonucu kontrol et
        if bannerheld_de_found or bannerheld_at_found:
            logger.info("En az bir işlem başarıyla gerçekleştirildi")
            return True
        else:
            logger.error("Hiçbir işlem gerçekleştirilemedi")
            return False
            
    except Exception as e:
        logger.error(f"Website seçim işlemleri sırasında hata: {str(e)}")
        return False

def configure_delivery_section(driver):
    """
    Delivery bölümünü bulup tıklar ve gerekli ayarları yapar:
    1. Disable pickup -> YES
    2. VPE -> 9999
    3. VPE Type -> Stock
    4. The product can be packaged in combination -> Fix Products
    """
    try:
        logger.info("Delivery bölümü ayarlanıyor...")
        
        
        # 1. "Delivery" başlığını bul ve tıkla
        delivery_found = False
        scroll_attempts = 0
        max_scroll_attempts = 20  # Maksimum kaydırma denemesi
        
        while not delivery_found and scroll_attempts < max_scroll_attempts:
            # Şu anki sayfada "Delivery" başlığını aramaya çalış
            try:
                # Görünür elementleri arama
                delivery_elements = driver.find_elements(By.XPATH, "//div[contains(text(), 'Delivery')] | //span[contains(text(), 'Delivery')]")
                
                # Görünür bir element bulundu mu kontrol et
                for element in delivery_elements:
                    if element.is_displayed():
                        # Element görünür, kaydır ve tıkla
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)  # Kaydırma sonrası bekle
                    
                        
                        # Elemente tıkla
                        driver.execute_script("arguments[0].click();", element)
                        logger.info("Delivery başlığına tıklandı")
                        
                        delivery_found = True
                        # Bir an bekleme - açılması için
                       
                        break
                
                if not delivery_found:
                    # Element bulunamadı, sayfayı biraz daha aşağı kaydır
                    driver.execute_script("window.scrollBy(0, 200);")  # 200px aşağı kaydır
                    scroll_attempts += 1
                    logger.info(f"Sayfayı aşağı kaydırıyorum - deneme {scroll_attempts}/{max_scroll_attempts}")
            
            except Exception as find_e:
                logger.warning(f"Element arama hatası: {str(find_e)}")
                # Sayfayı biraz daha aşağı kaydır ve tekrar dene
                driver.execute_script("window.scrollBy(0, 200);")
                scroll_attempts += 1
        
        # Eğer "Delivery" bulunamadıysa
        if not delivery_found:
            logger.error("Delivery başlığı bulunamadı, son bir JavaScript denemesi yapılıyor...")
            
            # Son bir JavaScript denemesi
            result = driver.execute_script("""
                // Sayfadaki tüm elementleri kontrol et
                var elements = document.querySelectorAll('*');
                
                for (var i = 0; i < elements.length; i++) {
                    // "Delivery" içeren bir element ara
                    if (elements[i].textContent && elements[i].textContent.trim() === 'Delivery') {
                        // Elementi görünür yap
                        elements[i].scrollIntoView({block: 'center'});
                        
                        // 500ms bekle
                        setTimeout(function() {}, 500);
                        
                        // Elementi tıkla
                        elements[i].click();
                        return "Delivery element bulundu ve tıklandı";
                    }
                }
                
                return "Delivery elementi bulunamadı";
            """)
            
            logger.info(f"JavaScript sonucu: {result}")
            delivery_found = "bulunamadı" not in result
            
    
            
            # Bölümün açılması için biraz bekle
           
        
        # Eğer hala bulunamadıysa, işlemi iptal et
        if not delivery_found:
            logger.error("Delivery bölümü bulunamadı, işleme devam edilemiyor.")
            return False
        
        
        # 2. Disable pickup -> YES
        logger.info("Disable pickup ayarlanıyor: YES")
        disable_pickup_result = set_disable_pickup(driver, True)
        
        if not disable_pickup_result:
            logger.warning("Disable pickup ayarlanamadı, devam ediliyor...")
        
        # 3. VPE -> 9999
        logger.info("VPE değeri ayarlanıyor: 9999")
        vpe_result = set_vpe_value(driver, "9999")
        
        if not vpe_result:
            logger.warning("VPE değeri ayarlanamadı, devam ediliyor...")
        
        # 4. VPE Type -> Stock
        logger.info("VPE Type ayarlanıyor: Stock")
        vpe_type_result = set_vpe_type(driver, "Stock")
        
        if not vpe_type_result:
            logger.warning("VPE Type ayarlanamadı, devam ediliyor...")
        
        # 5. The product can be packaged in combination -> Fix Products
        logger.info("Paketleme kombinasyonu ayarlanıyor: Fix Products")
        packaging_result = set_packaging_combination(driver, "Fix Products")
        
        if not packaging_result:
            logger.warning("Paketleme kombinasyonu ayarlanamadı, devam ediliyor...")
        
      
        
        # Başarı durumu - en az bir ayarlama yapılmışsa true döndür
        success = disable_pickup_result or vpe_result or vpe_type_result or packaging_result
        
        if success:
            logger.info("Delivery bölümü ayarları tamamlandı")
        else:
            logger.error("Delivery bölümü ayarlarının hiçbiri yapılamadı")
            
        return success
        
    except Exception as e:
        logger.error(f"Delivery bölümü ayarlanırken hata: {str(e)}")
        return False

def set_disable_pickup(driver, enable=True):
    """Disable pickup ayarını Yes veya No olarak ayarla"""
    try:
        logger.info(f"Disable pickup ayarlanıyor: {'Yes' if enable else 'No'}")
        
        # Disable pickup alanını bul
        result = driver.execute_script("""
            // "Disable pickup" etiketini bul
            var labels = Array.from(document.querySelectorAll('*')).filter(
                el => el.textContent && el.textContent.trim() === 'Disable pickup'
            );
            
            if (labels.length === 0) return {success: false, message: "Disable pickup etiketi bulunamadı"};
            
            var label = labels[0];
            
            // Etiketin görünür olduğundan emin ol
            label.scrollIntoView({block: 'center'});
            
            // Container elementi bul
            var container = label;
            while (container && !container.classList.contains('admin__field')) {
                container = container.parentElement;
                if (!container) return {success: false, message: "Container bulunamadı"};
            }
            
            // Toggle veya checkbox bul
            var toggle = container.querySelector('label.admin__actions-switch-label');
            var checkbox = container.querySelector('input[type="checkbox"]');
            
            if (!toggle && !checkbox) {
                return {success: false, message: "Toggle veya checkbox bulunamadı"};
            }
            
            // Mevcut durumu kontrol et
            var isEnabled = checkbox ? checkbox.checked : 
                         toggle.classList.contains('_active') || 
                         toggle.classList.contains('admin__actions-switch-checkbox-checked');
            
            // Sadece gerekirse değiştir
            if (isEnabled !== arguments[0]) {
                if (toggle) {
                    toggle.click();
                } else if (checkbox) {
                    checkbox.click();
                }
                
                if (checkbox) {
                    // Değişimi garantile
                    checkbox.checked = arguments[0];
                    checkbox.dispatchEvent(new Event('change', {bubbles: true}));
                }
                
                return {success: true, action: "değiştirildi", newState: arguments[0]};
            }
            
            return {success: true, action: "zaten doğru ayarda", currentState: isEnabled};
        """, enable)
        
        if isinstance(result, dict) and result.get('success'):
            logger.info(f"Disable pickup başarıyla ayarlandı: {result}")
            return True
        else:
            logger.warning(f"Disable pickup ayarlama başarısız: {result}")
            
            # Alternatif yaklaşım - direkt XPath ile bul
            try:
                disable_pickup_label = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Disable pickup')]"))
                )
                
                # Görünür olması için kaydır
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", disable_pickup_label)
                
                
                # Parent elementi bul
                parent = disable_pickup_label
                for _ in range(5):  # En fazla 5 seviye yukarı git
                    parent = parent.find_element(By.XPATH, "..")
                    if "admin__field" in parent.get_attribute("class"):
                        break
                
                # Switch elementi bul
                switch = parent.find_element(By.CSS_SELECTOR, ".admin__actions-switch-label")
                
                # Mevcut durumu kontrol et
                checkbox = parent.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                current_state = checkbox.is_selected()
                
                # Sadece gerekirse değiştir
                if current_state != enable:
                    driver.execute_script("arguments[0].click();", switch)
                    logger.info(f"Disable pickup başarıyla değiştirildi: {'Yes' if enable else 'No'}")
                else:
                    logger.info(f"Disable pickup zaten istenen durumda: {'Yes' if enable else 'No'}")
                
                return True
            except Exception as alt_e:
                logger.error(f"Alternatif yaklaşım da başarısız: {str(alt_e)}")
                return False
            
    except Exception as e:
        logger.error(f"Disable pickup ayarlanırken hata: {str(e)}")
        return False

def set_vpe_value(driver, value="9999"):
    """VPE alanına değer gir"""
    try:
        logger.info(f"VPE değeri giriliyor: {value}")
        
        # VPE alanını bul ve değeri gir
        result = driver.execute_script("""
            // "VPE" etiketini bul
            var labels = Array.from(document.querySelectorAll('*')).filter(
                el => el.textContent && el.textContent.trim() === 'VPE'
            );
            
            if (labels.length === 0) return {success: false, message: "VPE etiketi bulunamadı"};
            
            var label = labels[0];
            
            // Etiketin görünür olduğundan emin ol
            label.scrollIntoView({block: 'center'});
            
            // Container elementi bul
            var container = label;
            while (container && !container.classList.contains('admin__field')) {
                container = container.parentElement;
                if (!container) return {success: false, message: "Container bulunamadı"};
            }
            
            // Input alanını bul
            var input = container.querySelector('input');
            
            if (!input) {
                return {success: false, message: "Input alanı bulunamadı"};
            }
            
            // Değeri gir
            input.value = arguments[0];
            input.dispatchEvent(new Event('input', {bubbles: true}));
            input.dispatchEvent(new Event('change', {bubbles: true}));
            
            return {success: true, action: "değer girildi", value: arguments[0]};
        """, value)
        
        if isinstance(result, dict) and result.get('success'):
            logger.info(f"VPE değeri başarıyla girildi: {result}")
            return True
        else:
            logger.warning(f"VPE değeri girme başarısız: {result}")
            
            # Alternatif yaklaşım - direkt XPath ile bul
            try:
                vpe_label = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[text()='VPE']"))
                )
                
                # Görünür olması için kaydır
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vpe_label)
                
                
                # Parent elementi bul
                parent = vpe_label
                for _ in range(5):  # En fazla 5 seviye yukarı git
                    parent = parent.find_element(By.XPATH, "..")
                    if "admin__field" in parent.get_attribute("class"):
                        break
                
                # Input elementi bul
                input_elem = parent.find_element(By.CSS_SELECTOR, "input")
                
                # Değeri gir
                input_elem.clear()
                input_elem.send_keys(value)
                logger.info(f"VPE değeri başarıyla girildi: {value}")
                
                return True
            except Exception as alt_e:
                logger.error(f"Alternatif yaklaşım da başarısız: {str(alt_e)}")
                return False
            
    except Exception as e:
        logger.error(f"VPE değeri girilirken hata: {str(e)}")
        return False

def set_vpe_type(driver, type_value="Stock"):
    """VPE Type alanını ayarla"""
    try:
        logger.info(f"VPE Type ayarlanıyor: {type_value}")
        
        # VPE Type alanını bul ve değeri seç
        result = driver.execute_script("""
            // "VPE Type" etiketini bul
            var labels = Array.from(document.querySelectorAll('*')).filter(
                el => el.textContent && el.textContent.trim() === 'VPE Type'
            );
            
            if (labels.length === 0) return {success: false, message: "VPE Type etiketi bulunamadı"};
            
            var label = labels[0];
            
            // Etiketin görünür olduğundan emin ol
            label.scrollIntoView({block: 'center'});
            
            // Container elementi bul
            var container = label;
            while (container && !container.classList.contains('admin__field')) {
                container = container.parentElement;
                if (!container) return {success: false, message: "Container bulunamadı"};
            }
            
            // Select veya dropdown elementi bul
            var select = container.querySelector('select');
            var dropdown = container.querySelector('.admin__action-multiselect');
            
            if (select) {
                // Standard select elementi
                var options = Array.from(select.options);
                var targetOption = options.find(opt => opt.text.trim() === arguments[0]);
                
                if (targetOption) {
                    select.value = targetOption.value;
                    select.dispatchEvent(new Event('change', {bubbles: true}));
                    return {success: true, action: "standard select ile seçildi", value: arguments[0]};
                } else {
                    return {success: false, message: "Standard select'te hedef option bulunamadı"};
                }
            } else if (dropdown) {
                // Custom dropdown
                dropdown.click();
                
                // Dropdown açıldıktan sonra kısa bir bekleme
                setTimeout(function(){}, 200);
                
                // Dropdown option'ları içinde ara
                var options = document.querySelectorAll('.admin__action-multiselect-menu-inner-item span');
                var targetOption = Array.from(options).find(opt => opt.textContent.trim() === arguments[0]);
                
                if (targetOption) {
                    targetOption.click();
                    return {success: true, action: "custom dropdown ile seçildi", value: arguments[0]};
                } else {
                    return {success: false, message: "Custom dropdown'da hedef option bulunamadı"};
                }
            }
            
            return {success: false, message: "Select veya dropdown elementi bulunamadı"};
        """, type_value)
        
        if isinstance(result, dict) and result.get('success'):
            logger.info(f"VPE Type başarıyla ayarlandı: {result}")
            return True
        else:
            logger.warning(f"VPE Type ayarlama başarısız: {result}")
            
            # Alternatif yaklaşım - direkt XPath ile bul
            try:
                vpe_type_label = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[text()='VPE Type']"))
                )
                
                # Görünür olması için kaydır
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vpe_type_label)
                
                
                # Parent elementi bul
                parent = vpe_type_label
                for _ in range(5):  # En fazla 5 seviye yukarı git
                    parent = parent.find_element(By.XPATH, "..")
                    if "admin__field" in parent.get_attribute("class"):
                        break
                
                # Select elementi bul
                try:
                    select_elem = parent.find_element(By.CSS_SELECTOR, "select")
                    select = Select(select_elem)
                    select.select_by_visible_text(type_value)
                    logger.info(f"VPE Type başarıyla seçildi (select): {type_value}")
                    return True
                except:
                    # Custom dropdown olabilir
                    dropdown = parent.find_element(By.CSS_SELECTOR, ".admin__action-multiselect")
                    driver.execute_script("arguments[0].click();", dropdown)
                   
                    
                    # Açılan menüdeki seçeneği bul
                    option = driver.find_element(By.XPATH, f"//div[contains(@class, 'admin__action-multiselect-menu')]//span[text()='{type_value}']")
                    driver.execute_script("arguments[0].click();", option)
                    logger.info(f"VPE Type başarıyla seçildi (dropdown): {type_value}")
                    return True
                
            except Exception as alt_e:
                logger.error(f"Alternatif yaklaşım da başarısız: {str(alt_e)}")
                return False
            
    except Exception as e:
        logger.error(f"VPE Type ayarlanırken hata: {str(e)}")
        return False

def set_packaging_combination(driver, option_value="Fix Products"):
    """The product can be packaged in combination seçeneğini ayarla"""
    try:
        logger.info(f"Paketleme kombinasyonu ayarlanıyor: {option_value}")
        
        # Paketleme alanını bul ve değeri seç
        result = driver.execute_script("""
            // "The product can be packaged in combination" etiketini bul
            var labels = Array.from(document.querySelectorAll('*')).filter(
                el => el.textContent && el.textContent.includes('packaged in combination')
            );
            
            if (labels.length === 0) return {success: false, message: "Paketleme etiketi bulunamadı"};
            
            var label = labels[0];
            
            // Etiketin görünür olduğundan emin ol
            label.scrollIntoView({block: 'center'});
            
            // Container elementi bul
            var container = label;
            while (container && !container.classList.contains('admin__field')) {
                container = container.parentElement;
                if (!container) return {success: false, message: "Container bulunamadı"};
            }
            
            // Select elementi bul
            var select = container.querySelector('select');
            
            if (!select) {
                return {success: false, message: "Select elementi bulunamadı"};
            }
            
            // Seçeneği bul ve seç
            var options = Array.from(select.options);
            var targetOption = options.find(opt => opt.text.trim() === arguments[0]);
            
            if (targetOption) {
                select.value = targetOption.value;
                select.dispatchEvent(new Event('change', {bubbles: true}));
                return {success: true, action: "seçildi", value: arguments[0]};
            } else {
                return {success: false, message: "Hedef option bulunamadı", availableOptions: options.map(o => o.text)};
            }
        """, option_value)
        
        if isinstance(result, dict) and result.get('success'):
            logger.info(f"Paketleme kombinasyonu başarıyla ayarlandı: {result}")
            return True
        else:
            logger.warning(f"Paketleme kombinasyonu ayarlama başarısız: {result}")
            
            # Alternatif yaklaşım - direkt görünürdeki select elemanını kullanma
            try:
                # Görünür select elementlerini bul
                selects = driver.find_elements(By.CSS_SELECTOR, "select")
                
                for select_elem in selects:
                    if select_elem.is_displayed():
                        # Görünür elementin içeriğini kontrol et
                        options = select_elem.find_elements(By.CSS_SELECTOR, "option")
                        
                        for option in options:
                            if option.text.strip() == option_value:
                                # Doğru select bulundu, seçeneği seç
                                select = Select(select_elem)
                                select.select_by_visible_text(option_value)
                                logger.info(f"Paketleme kombinasyonu başarıyla seçildi: {option_value}")
                                return True
                
                logger.warning("Görünür select elementlerinde hedef seçenek bulunamadı")
                
                # Ekrandaki açık dropdown'ı bul ve Fix Products seçeneğini tıkla
                fix_products_option = driver.find_element(By.XPATH, "//li[text()='Fix Products']")
                driver.execute_script("arguments[0].click();", fix_products_option)
                logger.info(f"Fix Products seçeneği doğrudan tıklandı")
                return True
                
            except Exception as alt_e:
                logger.error(f"Alternatif yaklaşım da başarısız: {str(alt_e)}")
                return False
            
    except Exception as e:
        logger.error(f"Paketleme kombinasyonu ayarlanırken hata: {str(e)}")
        return False

def configure_web2print_settings(driver, product_part_options, product_options):
    """
    Web2Print ayarlarını yapar:
    1. Web2Print Settings sekmesini bulur ve tıklar
    2. General sekmesine tıklar
    3. Product Type -> "Upload + Personalization + Configuration" olarak ayarlar
    4. Matrix sekmesine tıklar
    5. Price Calculation -> Fixed olarak ayarlar
    6. Order of Product Part Options -> CSV'den gelen değer olarak ayarlar
    7. Order of Product Options -> CSV'den gelen değer olarak ayarlar

    """
    try:
        logger.info("Web2Print ayarları yapılıyor...")
        
        
        # 1. "Web2Print Settings" başlığını bul ve tıkla
        web2print_found = False
        scroll_attempts = 0
        max_scroll_attempts = 20  # Maksimum kaydırma denemesi
        
        # Mevcut konumdan başla, sayfanın başına gitme
        while not web2print_found and scroll_attempts < max_scroll_attempts:
            # Şu anki sayfada "Web2Print Settings" başlığını aramaya çalış
            try:
                # Görünür elementleri arama
                web2print_elements = driver.find_elements(By.XPATH, 
                    "//div[contains(text(), 'Web2Print Settings')] | //span[contains(text(), 'Web2Print Settings')]")
                
                # Görünür bir element bulundu mu kontrol et
                for element in web2print_elements:
                    if element.is_displayed():
                        # Element görünür, kaydır ve tıkla
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                        time.sleep(1)  # Kaydırma sonrası bekle
                        
                        # Elemente tıkla
                        driver.execute_script("arguments[0].click();", element)
                        logger.info("Web2Print Settings başlığına tıklandı")
                        
                        web2print_found = True
                        # Biraz bekle - açılması için
                        time.sleep(2)
                        break
                
                if not web2print_found:
                    # Element bulunamadı, sayfayı biraz daha aşağı kaydır
                    driver.execute_script("window.scrollBy(0, 200);")  # 200px aşağı kaydır
                    scroll_attempts += 1
                    time.sleep(0.5)  # Kaydırma sonrası bekle
                    logger.info(f"Sayfayı aşağı kaydırıyorum - deneme {scroll_attempts}/{max_scroll_attempts}")
            
            except Exception as find_e:
                logger.warning(f"Element arama hatası: {str(find_e)}")
                # Sayfayı biraz daha aşağı kaydır ve tekrar dene
                driver.execute_script("window.scrollBy(0, 200);")
                scroll_attempts += 1
                time.sleep(0.5)
        
        # Web2Print Settings bulunamadıysa, son bir JavaScript denemesi
        if not web2print_found:
            logger.warning("Web2Print Settings başlığı bulunamadı, JavaScript ile deniyorum...")
            
            result = driver.execute_script("""
                // Tüm elementleri kontrol et
                var elements = document.querySelectorAll('*');
                
                for (var i = 0; i < elements.length; i++) {
                    if (elements[i].textContent && elements[i].textContent.includes('Web2Print Settings')) {
                        // Elementi görünür yap
                        elements[i].scrollIntoView({block: 'center'});
                        
                        // 500ms bekle
                        setTimeout(function() {}, 500);
                        
                        // Elementi tıkla
                        elements[i].click();
                        return "Web2Print Settings bulundu ve tıklandı";
                    }
                }
                
                return "Web2Print Settings bulunamadı";
            """)
            
            logger.info(f"JavaScript sonucu: {result}")
            time.sleep(2)  # Sayfanın yüklenmesi için bekle
        
        # 2. General sekmesine tıkla
        try:
            general_tab = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='General']"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", general_tab)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", general_tab)
            logger.info("General sekmesine tıklandı")
            time.sleep(2)  # Sekmenin açılması için bekle
            
        
        except Exception as general_e:
            logger.warning(f"General sekmesi bulunamadı veya tıklanamadı: {str(general_e)}")
        
        try:
            # force_product_type_selection fonksiyonunu çağır
            product_type_result = update_product_type_like_producer(driver)
            if product_type_result:
                logger.info("Product Type başarıyla ayarlandı")
            else:
                logger.warning("Product Type ayarlanamadı, devam ediliyor...")
        except Exception as type_e:
            logger.warning(f"Product Type ayarlanamadı: {str(type_e)}")
        
        # 4. Matrix sekmesine tıkla
        try:
            matrix_tab = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Matrix']"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", matrix_tab)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", matrix_tab)
            logger.info("Matrix sekmesine tıklandı")
            time.sleep(2)  # Sekmenin açılması için bekle
            
           
        except Exception as matrix_e:
            logger.warning(f"Matrix sekmesi bulunamadı veya tıklanamadı: {str(matrix_e)}")
        
        # 5. Price Calculation -> Fixed olarak ayarla
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
        
        # 6. Order of Product Part Options -> CSV'den gelen değer olarak ayarla
        try:
            order_part_label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[text()='Order of Product Part Options']"))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", order_part_label)
            time.sleep(1)
            
            # Yakındaki input elementi bul
            input_script = """
                var label = arguments[0];
                var parent = label.closest('.admin__field');
                
                if (!parent) return null;
                
                var input = parent.querySelector('input');
                return input;
            """
            
            input_elem = driver.execute_script(input_script, order_part_label)
            
            if input_elem:
                # Input'a CSV'den gelen değeri gir
                driver.execute_script("""
                    var input = arguments[0];
                    
                    // Değeri temizle
                    input.value = '';
                    
                    // Yeni değeri gir
                    input.value = arguments[1];
                    input.dispatchEvent(new Event('input', {bubbles: true}));
                    input.dispatchEvent(new Event('change', {bubbles: true}));
                    
                    return "Order of Product Part Options değeri girildi";
                """, input_elem, product_part_options)
                
                logger.info(f"Order of Product Part Options -> {product_part_options} olarak ayarlandı")
                time.sleep(1)
            else:
                logger.warning("Order of Product Part Options input elementi bulunamadı")
                
          
        except Exception as order_part_e:
            logger.warning(f"Order of Product Part Options ayarlanamadı: {str(order_part_e)}")
        
        # 7. Order of Product Options -> CSV'den gelen değer olarak ayarla
        try:
            order_options_label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[text()='Order of Product Options']"))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", order_options_label)
            time.sleep(1)
            
            # Yakındaki input elementi bul
            input_script = """
                var label = arguments[0];
                var parent = label.closest('.admin__field');
                
                if (!parent) return null;
                
                var input = parent.querySelector('input');
                return input;
            """
            
            input_elem = driver.execute_script(input_script, order_options_label)
            
            if input_elem:
                # Input'a CSV'den gelen değeri gir
                driver.execute_script("""
                    var input = arguments[0];
                    
                    // Değeri temizle
                    input.value = '';
                    
                    // Yeni değeri gir
                    input.value = arguments[1];
                    input.dispatchEvent(new Event('input', {bubbles: true}));
                    input.dispatchEvent(new Event('change', {bubbles: true}));
                    
                    return "Order of Product Options değeri girildi";
                """, input_elem, product_options)
                
                logger.info(f"Order of Product Options -> {product_options} olarak ayarlandı")
                time.sleep(1)
            else:
                logger.warning("Order of Product Options input elementi bulunamadı")
                
          
        except Exception as order_opt_e:
            logger.warning(f"Order of Product Options ayarlanamadı: {str(order_opt_e)}")
    except Exception as e:
        logger.error(f"Web2Print ayarları yapılırken hata: {str(e)}")
        return False

def update_product_type_like_producer(driver, product_type="Upload + Personalization + Configuration"):
    """Update Product Type using the same successful approach that works for Producer"""
    try:
        logger.info(f"Updating Product Type to: {product_type} using producer-style approach")
        
        
        
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

def configure_personalization_tab(driver):
    """
    Personalization sekmesini bulup tıklayacak ve Open OTP/Designer in new page ayarını 
    Yes olarak yapacak bir fonksiyon
    """
    try:
        logger.info("Personalization sekmesini bulma ve ayarlama işlemi başlatılıyor")
        
    
        
        # 1. Personalization sekmesini bul ve tıkla - çoklu yöntem deneme
        personalization_found = False
        
        # İlk deneme: Doğrudan XPath ile bul
        try:
            personalization_tabs = driver.find_elements(By.XPATH, "//span[text()='Personalization']")
            
            for tab in personalization_tabs:
                if tab.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tab)
                    time.sleep(1.5)
                    driver.execute_script("arguments[0].click();", tab)
                    logger.info("Personalization sekmesi XPath ile bulundu ve tıklandı")
                    personalization_found = True
                    time.sleep(2)
                    break
        except Exception as xpath_e:
            logger.warning(f"XPath ile Personalization sekmesi bulunamadı: {str(xpath_e)}")
        
        # İkinci deneme: JavaScript ile tüm sekmeleri tara
        if not personalization_found:
            try:
                result = driver.execute_script("""
                    // Tüm sekmeleri bul
                    var tabs = document.querySelectorAll('.admin__page-nav-item, .ui-tabs-anchor, .admin__page-nav-link');
                    
                    // Önce tam eşleşme ara
                    for (var i = 0; i < tabs.length; i++) {
                        if (tabs[i].textContent && tabs[i].textContent.trim() === 'Personalization') {
                            tabs[i].scrollIntoView({block: 'center'});
                            setTimeout(function(){}, 500);
                            tabs[i].click();
                            return { success: true, method: "exact match" };
                        }
                    }
                    
                    // Kısmi eşleşme ara
                    for (var i = 0; i < tabs.length; i++) {
                        if (tabs[i].textContent && tabs[i].textContent.includes('Personalization')) {
                            tabs[i].scrollIntoView({block: 'center'});
                            setTimeout(function(){}, 500);
                            tabs[i].click();
                            return { success: true, method: "partial match" };
                        }
                    }
                    
                    // Tüm görünür elementleri kontrol et
                    var allElements = document.querySelectorAll('*');
                    for (var i = 0; i < allElements.length; i++) {
                        if (allElements[i].textContent && 
                            allElements[i].textContent.trim() === 'Personalization' && 
                            window.getComputedStyle(allElements[i]).display !== 'none') {
                            
                            allElements[i].scrollIntoView({block: 'center'});
                            setTimeout(function(){}, 500);
                            allElements[i].click();
                            return { success: true, method: "any element" };
                        }
                    }
                    
                    // 4. sekmeyi dene (genelde Personalization 4. sekmedir)
                    var tabContainers = document.querySelectorAll('.admin__page-nav-items, .ui-tabs');
                    for (var i = 0; i < tabContainers.length; i++) {
                        var tabItems = tabContainers[i].querySelectorAll('.admin__page-nav-item, .ui-tabs-anchor');
                        
                        if (tabItems.length >= 4) {
                            var fourthTab = tabItems[3];
                            fourthTab.scrollIntoView({block: 'center'});
                            setTimeout(function(){}, 500);
                            fourthTab.click();
                            return { success: true, method: "fourth tab" };
                        }
                    }
                    
                    return { success: false, message: "Personalization sekmesi bulunamadı" };
                """)
                
                if isinstance(result, dict) and result.get('success', True):
                    logger.info(f"Personalization sekmesi JavaScript ile bulundu ve tıklandı: {result.get('method')}")
                    personalization_found = True
                    time.sleep(2)
            except Exception as js_e:
                logger.warning(f"JavaScript ile Personalization sekmesi bulunamadı: {str(js_e)}")
        
        
        
        # 2. "Open OTP/Designer in new page" alanını bul ve YES olarak ayarla
        otp_toggle_success = False
        
        # İlk deneme: XPath ile bul
        try:
            otp_labels = driver.find_elements(By.XPATH, "//span[contains(text(), 'Open OTP/Designer in new page')]")
            
            for label in otp_labels:
                if label.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", label)
                    time.sleep(1.5)
                    
                    # Toggle/Switch elementini bul
                    parent = label
                    toggle = None
                    
                    # 5 seviye yukarı git
                    for _ in range(5):
                        try:
                            parent = parent.find_element(By.XPATH, "..")
                            toggle = parent.find_element(By.CSS_SELECTOR, ".admin__actions-switch-label")
                            break
                        except:
                            continue
                    
                    if toggle:
                        # Toggle'ı tıkla - yeşil (açık) duruma getir
                        driver.execute_script("arguments[0].click();", toggle)
                        logger.info("Open OTP/Designer toggle'ı bulundu ve tıklandı")
                        otp_toggle_success = True
                        time.sleep(1)
                        break
            
            if not otp_toggle_success:
                logger.warning("Open OTP/Designer toggle'ı XPath ile bulunamadı")
        except Exception as toggle_e:
            logger.warning(f"Open OTP/Designer toggle'ı bulunamadı: {str(toggle_e)}")
        
        # İkinci deneme: JavaScript ile daha esnek arama
        if not otp_toggle_success:
            try:
                result = driver.execute_script("""
                    // "Open OTP/Designer in new page" etiketini bul
                    var labels = Array.from(document.querySelectorAll('*')).filter(
                        el => el.textContent && el.textContent.includes('Open OTP/Designer')
                    );
                    
                    if (labels.length === 0) return { success: false, message: "OTP label bulunamadı" };
                    
                    var label = labels[0];
                    
                    // Field container'ı bul
                    var field = null;
                    var current = label;
                    
                    for (var i = 0; i < 10; i++) {
                        if (!current) break;
                        current = current.parentElement;
                        
                        if (current && current.classList && current.classList.contains('admin__field')) {
                            field = current;
                            break;
                        }
                    }
                    
                    if (!field) return { success: false, message: "Field container bulunamadı" };
                    
                    // Toggle elementini bul
                    var toggle = field.querySelector('.admin__actions-switch-label');
                    var checkbox = field.querySelector('input[type="checkbox"]');
                    
                    if (!toggle && !checkbox) return { success: false, message: "Toggle veya checkbox bulunamadı" };
                    
                    // Mevcut durumu kontrol et
                    var isEnabled = checkbox ? checkbox.checked : toggle.classList.contains('_active');
                    
                    // Eğer kapalıysa, aç
                    if (!isEnabled) {
                        if (toggle) toggle.click();
                        else if (checkbox) {
                            checkbox.checked = true;
                            checkbox.dispatchEvent(new Event('change', {bubbles: true}));
                        }
                        return { success: true, action: "enabled", message: "Toggle açıldı" };
                    }
                    
                    return { success: true, action: "already-enabled", message: "Toggle zaten açık" };
                """)
                
                if isinstance(result, dict) and result.get('success', False):
                    logger.info(f"Open OTP/Designer toggle'ı JavaScript ile ayarlandı: {result.get('action')}")
                    otp_toggle_success = True
                else:
                    logger.warning(f"JavaScript ile toggle ayarlanamadı: {result}")
            except Exception as js_toggle_e:
                logger.warning(f"JavaScript ile toggle ayarlanamadı: {str(js_toggle_e)}")
        
        
        # Başarı durumu
        if personalization_found or otp_toggle_success:
            logger.info("Personalization ayarları kısmen veya tamamen yapıldı")
            return True
        else:
            logger.error("Personalization ayarları yapılamadı")
            return False
    
    except Exception as e:
        logger.error(f"Personalization ayarları yapılırken beklenmeyen hata: {str(e)}")
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
        
        # "Print Files Options" kısmında seçenekleri bul ve seç
        try:
            logger.info("Trying to select Print Files Options")
            
            # Print Files Options alanını bul
            print_files_elements = driver.find_elements(By.XPATH, "//span[contains(text(), 'Print Files Options')] | //div[contains(text(), 'Print Files Options')]")
            
            if print_files_elements:
                for elem in print_files_elements:
                    if elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(1)
                        logger.info("Found Print Files Options area")
                        
                        # JavaScript ile seçenekleri bul ve seç
                        options_result = driver.execute_script("""
                            // Print Files Options elementinden başla
                            var printFilesElement = arguments[0];
                            var container = printFilesElement;
                            
                            // Gerekli seçenekler
                            var requiredOptions = [
                                "Upload from Account",
                                "Send to Graphic Department",
                                "Send by Email",
                                "Online Designer"
                            ];
                            
                            // Seçenekleri içeren container'ı bul (üst elementlere git)
                            for (var i = 0; i < 5; i++) {
                                container = container.parentElement;
                                if (!container) break;
                                
                                // Checkbox'ları veya seçenekleri ara
                                var options = container.querySelectorAll('input[type="checkbox"], select option, li, div[class*="option"]');
                                
                                if (options.length > 0) {
                                    // Her seçenek için kontrol et
                                    var selected = [];
                                    
                                    for (var j = 0; j < options.length; j++) {
                                        var option = options[j];
                                        var optionText = option.textContent || option.value || '';
                                        
                                        // Eğer etiket ile ilişkilendirilmiş bir checkbox ise
                                        if (option.type === 'checkbox') {
                                            var label = option.nextElementSibling;
                                            if (label && label.tagName.toLowerCase() === 'label') {
                                                optionText = label.textContent.trim();
                                            }
                                        }
                                        
                                        // İstenen seçeneklerden biri mi kontrol et
                                        if (requiredOptions.some(reqOption => optionText.includes(reqOption))) {
                                            // Checkbox ise ve seçili değilse seç
                                            if (option.type === 'checkbox' && !option.checked) {
                                                option.checked = true;
                                                option.dispatchEvent(new Event('change', {bubbles: true}));
                                                selected.push(optionText);
                                            } 
                                            // Select option ise ve seçili değilse seç
                                            else if (option.tagName && option.tagName.toLowerCase() === 'option' && !option.selected) {
                                                option.selected = true;
                                                option.dispatchEvent(new Event('change', {bubbles: true}));
                                                selected.push(optionText);
                                            }
                                            // Diğer türlü tıklanabilir bir element ise
                                            else if (option.tagName) {
                                                option.click();
                                                selected.push(optionText);
                                            }
                                        }
                                    }
                                    
                                    if (selected.length > 0) {
                                        return { success: true, selected: selected };
                                    }
                                }
                            }
                            
                            // Hiçbir seçenek bulunamadı veya seçilemedi
                            return { success: false, message: "Could not find or select options" };
                        """, elem)
                        
                        logger.info(f"JavaScript selection result: {options_result}")
                        
                        # JavaScript başarısız olduysa alternatif yaklaşım dene
                        if not isinstance(options_result, dict) or not options_result.get('success', False):
                            # İstenen seçenekleri doğrudan XPath ile bul ve tıkla
                            options_to_select = [
                                "Upload from Account",
                                "Send to Graphic Department",
                                "Send by Email",
                                "Online Designer"
                            ]
                            
                            for option in options_to_select:
                                try:
                                    option_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{option}')]")
                                    
                                    for option_elem in option_elements:
                                        if option_elem.is_displayed():
                                            # Önce checkbox kontrolü yap
                                            try:
                                                checkbox = option_elem.find_element(By.XPATH, "./preceding-sibling::input[@type='checkbox']")
                                                
                                                if not checkbox.is_selected():
                                                    driver.execute_script("arguments[0].click();", checkbox)
                                                    logger.info(f"Selected checkbox for: {option}")
                                            except:
                                                # Checkbox yoksa direkt tıkla
                                                driver.execute_script("arguments[0].click();", option_elem)
                                                logger.info(f"Clicked directly on: {option}")
                                            
                                            break
                                except Exception as option_e:
                                    logger.warning(f"Could not select {option}: {str(option_e)}")
                        
                        break
            else:
                logger.warning("Print Files Options area not found")

        except Exception as print_files_e:
            logger.warning(f"Error selecting Print Files Options: {str(print_files_e)}")

        driver.execute_script("window.scrollBy(0, 200);")
        time.sleep(1)
        
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

def configure_dpi_and_quantity_selectors(driver):
    """
    Minimum DPI, Use Custom Quantity Selector ve Use Fixed Price Quantity Selector
    ayarlarını yapılandıran fonksiyon
    """
    try:
        logger.info("Configuring Minimum DPI and Quantity Selectors")
        
      
        # 1. Minimum DPI ayarla (300)
        dpi_set = False
        try:
            # Minimum DPI labelını bul
            dpi_labels = driver.find_elements(By.XPATH, "//span[contains(text(), 'Minimum DPI')]")
            
            for dpi_label in dpi_labels:
                if dpi_label.is_displayed():
                    # Label'e scroll et
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dpi_label)
                    time.sleep(1)
                    
                    # Label'ın yanındaki input alanını bul
                    dpi_input = driver.execute_script("""
                        var label = arguments[0];
                        var field = null;
                        
                        // Container elementi bul
                        for (var i = 0; i < 5; i++) {
                            if (!label) break;
                            var parent = label.parentElement;
                            if (!parent) break;
                            
                            var input = parent.querySelector('input');
                            if (input) return input;
                            
                            label = parent;
                        }
                        
                        return null;
                    """, dpi_label)
                    
                    if dpi_input:
                        # Input'a değer gir
                        driver.execute_script("""
                            var input = arguments[0];
                            input.value = arguments[1];
                            input.dispatchEvent(new Event('input', {bubbles: true}));
                            input.dispatchEvent(new Event('change', {bubbles: true}));
                        """, dpi_input, "300")
                        
                        logger.info("Set Minimum DPI to 300")
                        dpi_set = True
                        break
        except Exception as dpi_e:
            logger.warning(f"Failed to set Minimum DPI: {str(dpi_e)}")
        
        # 2. Use Custom Quantity Selector -> YES
        custom_qty_set = False
        try:
            # Use Custom Quantity Selector labelını bul
            custom_qty_labels = driver.find_elements(By.XPATH, "//span[contains(text(), 'Use Custom Quantity Selector')]")
            
            for custom_qty_label in custom_qty_labels:
                if custom_qty_label.is_displayed():
                    # Label'e scroll et
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", custom_qty_label)
                    time.sleep(1)
                    
                    # Toggle'ı YES yap
                    toggle_result = driver.execute_script("""
                        var label = arguments[0];
                        var field = null;
                        
                        // Container elementi bul
                        for (var i = 0; i < 5; i++) {
                            if (!label) break;
                            label = label.parentElement;
                            if (!label) break;
                            
                            if (label.classList && label.classList.contains('admin__field')) {
                                field = label;
                                break;
                            }
                        }
                        
                        if (!field) return {success: false, message: "Field container not found"};
                        
                        // Toggle elementini bul
                        var toggle = field.querySelector('.admin__actions-switch-label');
                        var checkbox = field.querySelector('input[type="checkbox"]');
                        
                        if (!toggle && !checkbox) return {success: false, message: "Toggle/checkbox not found"};
                        
                        // Mevcut durumu kontrol et
                        var isEnabled = checkbox ? checkbox.checked : 
                                     toggle.classList.contains('_active');
                        
                        // Etkin değilse etkinleştir
                        if (!isEnabled) {
                            if (toggle) toggle.click();
                            else if (checkbox) {
                                checkbox.checked = true;
                                checkbox.dispatchEvent(new Event('change', {bubbles: true}));
                            }
                            return {success: true, action: "enabled"};
                        }
                        
                        return {success: true, action: "already-enabled"};
                    """, custom_qty_label)
                    
                    logger.info(f"Use Custom Quantity Selector result: {toggle_result}")
                    
                    if isinstance(toggle_result, dict) and toggle_result.get('success', False):
                        custom_qty_set = True
                    break
        except Exception as custom_qty_e:
            logger.warning(f"Failed to set Use Custom Quantity Selector: {str(custom_qty_e)}")
        
        # 3. Use Fixed Price Quantity Selector -> YES
        fixed_qty_set = False
        try:
            # Use Fixed Price Quantity Selector labelını bul
            fixed_qty_labels = driver.find_elements(By.XPATH, "//span[contains(text(), 'Use Fixed Price Quantity Selector')]")
            
            for fixed_qty_label in fixed_qty_labels:
                if fixed_qty_label.is_displayed():
                    # Label'e scroll et
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", fixed_qty_label)
                    time.sleep(1)
                    
                    # Toggle'ı YES yap
                    toggle_result = driver.execute_script("""
                        var label = arguments[0];
                        var field = null;
                        
                        // Container elementi bul
                        for (var i = 0; i < 5; i++) {
                            if (!label) break;
                            label = label.parentElement;
                            if (!label) break;
                            
                            if (label.classList && label.classList.contains('admin__field')) {
                                field = label;
                                break;
                            }
                        }
                        
                        if (!field) return {success: false, message: "Field container not found"};
                        
                        // Toggle elementini bul
                        var toggle = field.querySelector('.admin__actions-switch-label');
                        var checkbox = field.querySelector('input[type="checkbox"]');
                        
                        if (!toggle && !checkbox) return {success: false, message: "Toggle/checkbox not found"};
                        
                        // Mevcut durumu kontrol et
                        var isEnabled = checkbox ? checkbox.checked : 
                                     toggle.classList.contains('_active');
                        
                        // Etkin değilse etkinleştir
                        if (!isEnabled) {
                            if (toggle) toggle.click();
                            else if (checkbox) {
                                checkbox.checked = true;
                                checkbox.dispatchEvent(new Event('change', {bubbles: true}));
                            }
                            return {success: true, action: "enabled"};
                        }
                        
                        return {success: true, action: "already-enabled"};
                    """, fixed_qty_label)
                    
                    logger.info(f"Use Fixed Price Quantity Selector result: {toggle_result}")
                    
                    if isinstance(toggle_result, dict) and toggle_result.get('success', False):
                        fixed_qty_set = True
                    break
        except Exception as fixed_qty_e:
            logger.warning(f"Failed to set Use Fixed Price Quantity Selector: {str(fixed_qty_e)}")
        
      
        
        # En az bir ayar başarılı olduysa True dön
        return dpi_set or custom_qty_set or fixed_qty_set
    
    except Exception as e:
        logger.error(f"Failed to configure DPI and Quantity Selectors: {str(e)}")
        return False

def select_last_product_part(driver):
    """Print Product Setup bölümünde Product Parts listesinde en son öğeyi seç"""
    try:
        logger.info("Attempting to select the last option in Product Parts list")
        
    
        
        # Product Parts'ı bul (görseldeki kırmızı kutu içindeki alan)
        try:
            # Direkt XPath kullanarak select elementi bul
            product_parts_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                    "//span[text()='Product Parts']/ancestor::div[contains(@class, 'field')]//select | //label[contains(text(), 'Product Parts')]/following-sibling::div//select"))
            )
            
            # Görünür olması için sayfayı kaydır
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", product_parts_select)
            time.sleep(1)
            
            # Element bulundu mu loglama yap
            logger.info("Found Product Parts select element")
            
            # Select sınıfını oluştur
            select = Select(product_parts_select)
            
            # Seçenek sayısını al
            options = select.options
            option_count = len(options)
            
            if option_count > 0:
                # En son seçeneğin indexi
                last_index = option_count - 1
                
                # Seçimi temizle (varsa)
                select.deselect_all()
                
                # En son seçeneği seç
                select.select_by_index(last_index)
                logger.info(f"Selected last option (index {last_index}) from Product Parts list")
            else:
                logger.warning("No options found in Product Parts dropdown")
                return False
                
        except Exception as e:
            logger.warning(f"Standard approach failed: {str(e)}")
            
            # JavaScript alternatif yaklaşım
            try:
                result = driver.execute_script("""
                    // Product Parts alanını bul
                    var productPartsLabel = Array.from(document.querySelectorAll('*')).find(
                        el => el.textContent === 'Product Parts'
                    );
                    
                    if (!productPartsLabel) {
                        productPartsLabel = Array.from(document.querySelectorAll('label, span')).find(
                            el => el.textContent.includes('Product Parts')
                        );
                    }
                    
                    if (!productPartsLabel) {
                        return {success: false, message: "Product Parts label not found"};
                    }
                    
                    // Label'dan select elementine git
                    var select = null;
                    
                    // Önce parent elementi kontrol et
                    var parent = productPartsLabel.closest('div[class*="field"]');
                    if (parent) {
                        select = parent.querySelector('select');
                    }
                    
                    // Parent'ta bulamazsak, label'a yakın select'i bul
                    if (!select) {
                        // Tüm select elementlerini bul
                        var allSelects = document.querySelectorAll('select');
                        var labelRect = productPartsLabel.getBoundingClientRect();
                        
                        // En yakın select'i bul
                        var closestSelect = null;
                        var minDistance = Number.MAX_VALUE;
                        
                        for (var i = 0; i < allSelects.length; i++) {
                            var selectRect = allSelects[i].getBoundingClientRect();
                            var distance = Math.sqrt(
                                Math.pow(selectRect.left - labelRect.left, 2) + 
                                Math.pow(selectRect.top - labelRect.top, 2)
                            );
                            
                            if (distance < minDistance) {
                                minDistance = distance;
                                closestSelect = allSelects[i];
                            }
                        }
                        
                        if (closestSelect) {
                            select = closestSelect;
                        }
                    }
                    
                    if (!select) {
                        return {success: false, message: "Could not find select element"};
                    }
                    
                    // Seçenek sayısını al
                    var options = select.options;
                    var optionCount = options.length;
                    
                    if (optionCount === 0) {
                        return {success: false, message: "Select has no options"};
                    }
                    
                    // İndex değerleri sayla
                    var optionTexts = [];
                    for (var i = 0; i < optionCount; i++) {
                        optionTexts.push(options[i].text);
                    }
                    
                    // Tüm seçenekleri temizle
                    for (var i = 0; i < optionCount; i++) {
                        options[i].selected = false;
                    }
                    
                    // En son seçeneği seç
                    var lastIndex = optionCount - 1;
                    options[lastIndex].selected = true;
                    
                    // Change olayını tetikle
                    select.dispatchEvent(new Event('change', {bubbles: true}));
                    
                    return {
                        success: true, 
                        message: "Selected last option", 
                        index: lastIndex,
                        text: options[lastIndex].text,
                        allOptions: optionTexts
                    };
                """)
                
                logger.info(f"JavaScript result: {result}")
                
                if isinstance(result, dict) and result.get('success', False):
                    logger.info(f"Successfully selected last option using JavaScript: {result.get('text')}")
                else:
                    logger.warning("JavaScript approach also failed")
                    return False
            except Exception as js_e:
                logger.error(f"JavaScript approach failed: {str(js_e)}")
                return False
        
        
        return True
    except Exception as e:
        logger.error(f"Failed to select last option in Product Parts: {str(e)}")
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
                                input.value = 'Banner';
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
                                    if (select.options[i].text === 'Banner') {
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
                                    allElements[i].textContent.trim() === 'Banner' &&
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
    
def save_product_printnet(driver):
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

def configure_all_settings_printnet(driver, product_name, sku, product_part_options, product_options):
    """Fill all required settings in PrintNet"""
    try:
        # 1. navigation
        if not navigate_to_add_product_printnet(driver):
            logger.warning("Failed to set 'Is Meter Price' in PrintNet")
            # Continue anyway

        # 2. Temel bilgileri doldur (name, sku, price, quantity)
        if not fill_basic_info_printnet(driver, product_name, sku):
            logger.error(f"Failed to fill basic info in PrintNet for product: {product_name}")
            return False
        
        if not set_visibility_to_catalog(driver):
            logger.error(f"Failed catalog")
            return False

        # 3. Is Meter Price ayarını Yes yap
        if not set_meter_price_printnet(driver, is_meter_price=True):
            logger.warning("Failed to set 'Is Meter Price' in PrintNet")
            # Continue anyway
            
        # 4. Set Product in Websites
        if not process_printnet(driver):
            logger.error("Failed to set 'Product in Websites' in Printnet")
            #Continue anyway 
            
        # 5. Delivery bölümü ayarları
        if not configure_delivery_section(driver):
            logger.warning("Failed to configure Delivery section")
            # Continue anyway
            
        # 6. Web2Print ve Personalization ayarları - CSV'den değerleri geçir
        if not configure_web2print_settings(driver, product_part_options, product_options):
            logger.warning("Failed to configure Web2Print settings")
            # Continue anyway
        if not configure_personalization_tab(driver):
            logger.warning("Failed to configure Web2Print settings")
            # Continue anyway
        if not select_production_days_multiselect(driver):
            logger.warning("Failed to configure Web2Print settings")
            # Continue anyway
        if not configure_dpi_and_quantity_selectors(driver):
            logger.warning("Failed to configure Web2Print settings")
            # Continue anyway    
        if not configure_print_product_setup_with_ctrl(driver):
            logger.warning("Failed to configure Web2Print settings")
            # Continue anyway 
        if not configure_additional_settings_fix(driver):
            logger.warning("Failed to configure Web2Print settings")
            # Continue anyway 
        logger.info(f"Successfully configured all settings in PrintNet for product: {product_name}")
        return True
    except Exception as e:
        logger.error(f"Failed to configure all settings in PrintNet: {str(e)}")
        return False

def main():
    """Main function to run the PrintNet automation for the first product"""
    # Check if CSV file exists
    if not os.path.exists(CSV_FILE_PATH):
        logger.error(f"CSV file not found: {CSV_FILE_PATH}")
        return
    
    # Get user credentials
    printnet_username = username
    printnet_password = Password    
    # Setup Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    
    # Initialize Chrome driver
    driver = None
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        logger.info("Chrome WebDriver initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize Chrome WebDriver: {str(e)}")
        return
    
    try:
        # Login to PrintNet
        if not login_to_printnet(driver, printnet_username, printnet_password):
            logger.error("Login to PrintNet failed. Cannot proceed.")
            return

        logger.info("Successfully logged in to PrintNet")
        
        # Open and read the CSV file to get the first product
        try:
            with open(CSV_FILE_PATH, 'r', encoding='utf-8') as file:
                csv_reader = csv.reader(file)
                
                # Skip header rows (first 6 rows)
                for _ in range(6):
                    next(csv_reader, None)
                
                # Read only the first product
                first_product_row = next(csv_reader, None)
                
                if not first_product_row:
                    logger.warning("No valid product data found in CSV")
                    return

                try:
                    # Extract product data from CSV
                    product_name = first_product_row[CSV_MAPPINGS["product_name"]]
                    sku = first_product_row[CSV_MAPPINGS["sku"]]
                    # Extract the missing parameters
                    product_part_options = first_product_row[CSV_MAPPINGS["product_part_options"]]
                    product_options = first_product_row[CSV_MAPPINGS["product_options"]]
                    
                    logger.info(f"Processing first product: {product_name} (SKU: {sku})")
                    
                    # Configure all settings with the product data
                    if not configure_all_settings_printnet(driver, product_name, sku, product_part_options, product_options):
                        logger.error(f"Failed to configure product: {product_name}")
                        return

                    logger.info(f"Successfully configured product: {product_name}")
                    
                    # Save the product
                    if not save_product_printnet(driver):
                        logger.error(f"Failed to save product: {product_name}")
                        return

                    logger.info(f"Successfully saved product: {product_name}")

                except IndexError:
                    logger.error("Product row has invalid format or missing data")
                except Exception as product_e:
                    logger.error(f"Error processing product: {str(product_e)}")

        except FileNotFoundError:
            logger.error(f"CSV file not found: {CSV_FILE_PATH}")
        except PermissionError:
            logger.error(f"Permission denied when accessing CSV file: {CSV_FILE_PATH}")
        except Exception as csv_e:
            logger.error(f"Error processing CSV file: {str(csv_e)}")

    except Exception as e:
        logger.error(f"Unexpected error during PrintNet automation: {str(e)}")

    finally:
        # Close the browser
        if driver:
            try:
                driver.quit()
                logger.info("Browser closed")
            except Exception as close_e:
                logger.error(f"Error closing browser: {str(close_e)}")

if __name__ == "__main__":
    main()