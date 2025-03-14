import time
import pyautogui
import requests
from bs4 import BeautifulSoup
import webbrowser
import pygetwindow as gw
import logging
import os
import csv
from alive_progress import alive_bar
from urllib.parse import urljoin, urlparse
import pyperclip

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraper_log.txt'),
        logging.StreamHandler()
    ]
)

chrome_path = r"C:\Users\ZnK\AppData\Local\Google\Chrome SxS\Application\chrome.exe"
visited_urls_file = 'visited_urls.txt'
search_terms = [
    "plumbers new york",
    "electricians manhattan",
    "dentists london",
    "auto repair florida"
]
results_csv = 'website_results.csv'
max_pages = 3
chatbot_markers = [
    'chat-widget', 'livechat', 'chat-bot', 'chatbot', 'live-chat', 
    'bot-widget', 'zendesk', 'intercom', 'drift', 'freshchat',
    'crisp', 'tawk', 'liveperson', 'olark', 'userlike',
    'chat_widget', 'chat_container', 'chat_bubble', 'chat_frame'
]

def get_base_url(url):
    parsed_uri = urlparse(url)
    return f'{parsed_uri.scheme}://{parsed_uri.netloc}'

def load_visited_urls(file_path):
    visited_urls = set()
    if os.path.exists(file_path):
        with open(file_path, 'r') as file:
            for line in file:
                url = line.strip()
                if url:
                    # Add both the full URL and its base domain to avoid duplicates
                    visited_urls.add(url)
                    base_url = get_base_url(url)
                    if base_url:
                        visited_urls.add(base_url)
    logging.info(f"Loaded {len(visited_urls)} visited URLs from {file_path}")
    return visited_urls

def save_visited_urls(file_path, visited_urls):
    # Ensure we don't lose existing URLs when saving
    existing_urls = set()
    if os.path.exists(file_path):
        try:
            with open(file_path, 'r') as file:
                for line in file:
                    url = line.strip()
                    if url:
                        existing_urls.add(url)
        except Exception as e:
            logging.error(f"Error reading existing URLs: {str(e)}")
    
    # Combine with new URLs
    all_urls = existing_urls.union(visited_urls)
    
    # Write all URLs back to file with error handling and immediate flush
    try:
        with open(file_path, 'w') as file:
            for url in all_urls:
                file.write(url + '\n')
            file.flush()  # Force write to disk immediately
            os.fsync(file.fileno())  # Ensure it's written to the physical disk
        
        logging.info(f"Saved {len(all_urls)} URLs to {file_path}")
    except Exception as e:
        logging.error(f"Error saving URLs to file: {str(e)}")
        # Try alternative approach with append mode if write mode fails
        try:
            with open(file_path, 'a') as file:
                for url in visited_urls:
                    if url not in existing_urls:
                        file.write(url + '\n')
                file.flush()
                os.fsync(file.fileno())
            logging.info(f"Saved new URLs using append mode")
        except Exception as append_error:
            logging.error(f"Error saving URLs in append mode: {str(append_error)}")

def initialize_csv():
    if not os.path.exists(results_csv):
        with open(results_csv, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Search Term', 'URL', 'Title', 'Description', 'Has Chatbot', 'Phone', 'Email', 'Contact URL', 'Location'])

def save_to_csv(search_term, business_name, url, has_chatbot, phone, email, contact_url, address, description, location):
    # Format has_chatbot as YES/NO instead of True/False
    has_chatbot_formatted = "YES" if has_chatbot else "NO"
    
    # Format phone number to ensure it has country code if available
    formatted_phone = phone
    if phone and not phone.startswith("+"):
        formatted_phone = "+1 " + phone
    
    # Clean up location data - remove 'Locations' prefix if present
    cleaned_location = location
    if cleaned_location and cleaned_location.lower().startswith("locations"):
        cleaned_location = cleaned_location[9:].strip()  # Remove 'Locations' and any leading whitespace
    
    try:
        with open(results_csv, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            # Ensure the order matches the desired format: Search Term,URL,Title,Description,Has Chatbot,Phone,Email,Contact URL,Location
            writer.writerow([search_term, url, business_name, description, has_chatbot_formatted, formatted_phone, email, contact_url, cleaned_location])
            file.flush()  # Force write to disk immediately
            os.fsync(file.fileno())  # Ensure it's written to the physical disk
        logging.info(f"Saved data for {business_name} - {url}")
    except Exception as e:
        logging.error(f"Error saving data to CSV for {url}: {str(e)}")
        # Try alternative approach with direct file writing if CSV writer fails
        try:
            with open(results_csv, 'a', encoding='utf-8') as file:
                row_data = f'"{search_term}","{url}","{business_name}","{description}","{has_chatbot_formatted}","{formatted_phone}","{email}","{contact_url}","{cleaned_location}"\n'
                file.write(row_data)
                file.flush()
                os.fsync(file.fileno())
            logging.info(f"Saved data using alternative method for {business_name}")
        except Exception as alt_error:
            logging.error(f"Alternative CSV save method failed: {str(alt_error)}")

def locate_button(image_path, timeout=10, region=None, confidence=0.8):
    start_time = time.time()
    # Try with multiple monitor configurations
    screen_width, screen_height = pyautogui.size()
    
    # Store original image path for logging
    original_image_path = image_path
    
    # Check if image file exists
    if not os.path.exists(image_path):
        logging.error(f"Image file not found: {image_path}")
        return None
    
    while time.time() - start_time < timeout:
        try:
            # Try different confidence levels if initial attempt fails
            confidence_levels = [confidence, 0.7, 0.6, 0.5, 0.4, 0.3]
            for conf in confidence_levels:
                # Try with the specified region first
                if region:
                    button = pyautogui.locateOnScreen(image_path, confidence=conf, region=region)
                    if button:
                        logging.info(f"Button found: {original_image_path} with confidence {conf}")
                        return button
                
                # If no region specified or not found in region, try full screen
                button = pyautogui.locateOnScreen(image_path, confidence=conf)
                if button:
                    logging.info(f"Button found: {original_image_path}")
                    return button
                
                # Small delay between retries
                time.sleep(0.2)
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as e:
            logging.warning(f"Image search error: {str(e)}")
            # Continue trying with other confidence levels
        
        time.sleep(0.5)
    
    logging.warning(f"Button not found: {original_image_path} after trying multiple confidence levels")
    return None

def open_search_engine(search_term):
    """Open Google search with the specified search term"""
    search_url = f"https://www.google.com/search?q={search_term.replace(' ', '+')}"
    webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
    browser = webbrowser.get('chrome')
    browser.open_new(search_url)
    
    time.sleep(5)
    chrome_windows = gw.getWindowsWithTitle("Google Chrome")
    if not chrome_windows:
        logging.error("Chrome window not found")
        return None
    
    main_window = chrome_windows[0]
    main_window.activate()
    
    screen_width, screen_height = pyautogui.size()
    main_window.resizeTo(screen_width // 2, screen_height)
    main_window.moveTo(screen_width // 2, 0)
    
    return main_window

def click_places_tab():
    logging.info("Looking for Places tab")
    time.sleep(7)
    
    # Try to find the Places tab with different confidence levels
    places_tab = locate_button("img/places.png", timeout=10, confidence=0.7)
    if places_tab:
        center_x, center_y = pyautogui.center(places_tab)
        logging.info(f"Places tab found at position: ({center_x}, {center_y})")
        pyautogui.click(places_tab)
        time.sleep(3)
        return True
    
    # If Places tab not found, try fallback mechanisms
    try:
        # Get screen dimensions for all monitors
        screen_width, screen_height = pyautogui.size()
        
        # Try multiple potential positions where Places tab might be
        fallback_positions = [
            (screen_width // 4, 200),  # Standard position
            (480, 200),               # Fixed position that worked before
            (screen_width // 4, 150), # Slightly higher
            (screen_width // 4, 250)  # Slightly lower
        ]
        
        for pos_x, pos_y in fallback_positions:
            logging.info(f"Falling back to approximate position: ({pos_x}, {pos_y})")
            pyautogui.click(pos_x, pos_y)
            time.sleep(2)
            
            # Check if we can find the web.png button after clicking, which would indicate
            # that we successfully clicked the Places tab
            web_button = locate_button("img/webss.png", timeout=3, confidence=0.6)
            if web_button:
                logging.info("Places tab click successful (verified by finding web button)")
                return True
        
        # If all fallback positions fail, use the most reliable one as last resort
        logging.warning("All fallback positions failed, using last resort position")
        pyautogui.click(480, 200)
        time.sleep(3)
        return True
    except Exception as e:
        logging.error(f"Error clicking Places tab: {str(e)}")
        return False

# Add new configuration variables
scroll_step = 300  # Pixels to scroll per step
result_offset = 80  # Vertical distance between results

def count_website_buttons(timeout=5):
    """Count all website buttons visible on the current page with progressive scrolling"""
    screen_width, screen_height = pyautogui.size()
    all_buttons = []
    y_position = 300  # Start from top of results
    max_scroll_attempts = 7  # Increased maximum number of scroll attempts
    scroll_attempts = 0
    last_button_count = 0
    
    # First scan the initially visible area
    while y_position < screen_height - 100:
        region = (0, y_position, screen_width, 300)  # Search in a window of 300px height
        try:
            # Try multiple confidence levels for better detection
            for confidence in [0.8, 0.7, 0.6, 0.5]:
                button = pyautogui.locateOnScreen("img/webss.png", confidence=confidence, region=region)
                if button:
                    center_x, center_y = pyautogui.center(button)
                    # Check if this button is already in our list (avoid duplicates)
                    if not any(abs(center_y - y) < 20 and abs(center_x - x) < 20 for x, y in all_buttons):
                        all_buttons.append((center_x, center_y))
                        logging.info(f"Found website button at ({center_x}, {center_y}) with confidence {confidence}")
                    y_position = center_y + 50  # Move past this button
                    break
            else:  # No button found at any confidence level
                y_position += 200  # Move down if no button found
        except Exception as e:
            logging.warning(f"Error during button counting: {str(e)}")
            y_position += 200  # Move down on error
    
    # Now progressively scroll and scan for more buttons
    while scroll_attempts < max_scroll_attempts:
        # Remember how many buttons we had before scrolling
        last_button_count = len(all_buttons)
        
        # Scroll down to reveal more results
        pyautogui.scroll(-700)  # Scroll down
        time.sleep(1.5)  # Increased wait time for scroll to complete
        
        # Scan the newly revealed area
        y_position = 300  # Reset scan position after scroll
        while y_position < screen_height - 100:
            region = (0, y_position, screen_width, 300)
            try:
                # Try multiple confidence levels for better detection
                for confidence in [0.8, 0.7, 0.6, 0.5]:
                    button = pyautogui.locateOnScreen("img/webss.png", confidence=confidence, region=region)
                    if button:
                        center_x, center_y = pyautogui.center(button)
                        # Check if this button is already in our list (avoid duplicates)
                        if not any(abs(center_y - y) < 20 and abs(center_x - x) < 20 for x, y in all_buttons):
                            all_buttons.append((center_x, center_y))
                            logging.info(f"Found website button at ({center_x}, {center_y}) with confidence {confidence} after scroll")
                        y_position = center_y + 50
                        break
                else:  # No button found at any confidence level
                    y_position += 200
            except Exception as e:
                logging.warning(f"Error during button counting after scroll: {str(e)}")
                y_position += 200
        
        # If we didn't find any new buttons after scrolling, try again
        if len(all_buttons) == last_button_count:
            scroll_attempts += 1
        else:
            # Reset scroll attempts if we found new buttons
            scroll_attempts = 0
    
    # Scroll back to top for consistent starting position
    for _ in range(max_scroll_attempts):
        pyautogui.scroll(700)  # Scroll up
        time.sleep(0.5)
    
    logging.info(f"Found {len(all_buttons)} website buttons on current page")
    return all_buttons

def click_website_button(timeout=5, previous_y=None, button_positions=None, button_index=0):
    """Click website button with position tracking"""
    screen_width, screen_height = pyautogui.size()
    
    # If we have pre-counted button positions, use those
    if button_positions and button_index < len(button_positions):
        try:
            center_x, center_y = button_positions[button_index]
            logging.info(f"Clicking pre-counted button {button_index+1}/{len(button_positions)} at position ({center_x}, {center_y})")
            
            # Validate coordinates are within screen bounds
            if center_x < 0 or center_x > screen_width or center_y < 0 or center_y > screen_height:
                logging.warning(f"Button coordinates ({center_x}, {center_y}) are outside screen bounds ({screen_width}x{screen_height})")
                return None, False
            
            # Ensure the button is visible - scroll if needed
            if center_y > screen_height - 200:
                # Need to scroll down to make button visible
                scroll_amount = min(center_y - (screen_height // 2), 700)  # Don't scroll too much at once
                pyautogui.scroll(-scroll_amount)  # Negative to scroll down
                time.sleep(1.5)  # Wait for scroll to complete
                
                # Recalculate button position after scroll
                try:
                    # Try to locate the button again after scrolling
                    region = (center_x - 100, center_y - 300, 200, 600)  # Search around expected position
                    button = pyautogui.locateOnScreen("img/webss.png", confidence=0.6, region=region)
                    if button:
                        center_x, center_y = pyautogui.center(button)
                        logging.info(f"Recalculated button position after scroll: ({center_x}, {center_y})")
                except Exception as scroll_error:
                    logging.warning(f"Error recalculating position after scroll: {str(scroll_error)}")
                    # Continue with original coordinates
            
            # Open in new tab
            pyautogui.moveTo(center_x, center_y)
            time.sleep(0.8)  # Increased delay to ensure cursor is positioned
            
            # Try to click with error handling
            try:
                pyautogui.keyDown('ctrl')
                pyautogui.click()
                pyautogui.keyUp('ctrl')
            except Exception as click_error:
                logging.error(f"Error during ctrl+click: {str(click_error)}")
                # Try alternative approach
                try:
                    pyautogui.click(button='right')
                    time.sleep(0.8)
                    pyautogui.press('t')  # 'Open in new tab' option
                except Exception as alt_error:
                    logging.error(f"Alternative click method failed: {str(alt_error)}")
                    return None, False
            
            # Move cursor to avoid hover effects
            pyautogui.moveTo(center_x, center_y + result_offset)
            time.sleep(3.5)  # Increased wait time for tab to open
            return center_y + result_offset, True  # Return position and success flag
        except Exception as e:
            logging.error(f"Error using pre-counted button {button_index}: {str(e)}")
            # Fall through to fallback method
    
    # Fallback to old method if no pre-counted buttons or index out of range
    min_region_height = 200  # Increased minimum search area height
    
    # Define regions to search in
    if previous_y is not None:
        # Ensure we leave enough search area
        region_height = max(screen_height - previous_y, min_region_height)
        region = (0, previous_y, screen_width, region_height)
    else:
        region = (0, 300, screen_width, screen_height - 300)

    # Try to find the web button in the specified region
    website_button = None
    try:
        website_button = locate_button("img/webss.png", timeout=timeout, region=region)
    except Exception as e:
        logging.warning(f"Error locating button in region: {str(e)}")
        # Continue to fallback approach
    
    # If not found in region, try full screen
    if not website_button:
        try:
            logging.info("Trying to find web button in full screen")
            website_button = locate_button("img/webss.png", timeout=timeout)
        except Exception as e:
            logging.warning(f"Error locating button in full screen: {str(e)}")
            return None, False

    if website_button:
        try:
            center_x, center_y = pyautogui.center(website_button)
            
            # Open in new tab
            pyautogui.moveTo(center_x, center_y)
            time.sleep(0.5)  # Small delay to ensure cursor is positioned
            
            # Try to click with error handling
            try:
                pyautogui.keyDown('ctrl')
                pyautogui.click()
                pyautogui.keyUp('ctrl')
            except Exception as click_error:
                logging.error(f"Error during ctrl+click: {str(click_error)}")
                # Try alternative approach
                try:
                    pyautogui.click(button='right')
                    time.sleep(0.5)
                    pyautogui.press('t')  # 'Open in new tab' option
                except Exception as alt_error:
                    logging.error(f"Alternative click method failed: {str(alt_error)}")
                    return None, False
            
            # Move cursor to avoid hover effects
            pyautogui.moveTo(center_x, center_y + result_offset)
            time.sleep(3)
            return center_y + result_offset, True  # Return last clicked position and success flag
        except Exception as e:
            logging.error(f"Error during button click: {str(e)}")
            return None, False
    else:
        logging.warning("Button not found: img/webss.png")
        return None, False

def process_business_listing(main_window, visited_urls, search_term, last_y_position, button_positions=None, button_index=0):
    try:
        # Store the original window title to help identify the main results tab
        original_title = None
        try:
            main_window.activate()
            time.sleep(2)  # Increased delay for reliable window activation
            original_title = main_window.title
            logging.info(f"Original tab title: {original_title}")
        except Exception as window_error:
            logging.error(f"Window activation error: {str(window_error)}")
            # Try to recover by finding a Chrome window
            chrome_windows = gw.getWindowsWithTitle("Google Chrome")
            if chrome_windows:
                main_window = chrome_windows[0]
                try:
                    main_window.activate()
                    time.sleep(2)  # Increased delay
                    original_title = main_window.title
                    logging.info(f"Recovered tab title: {original_title}")
                except Exception:
                    logging.error("Failed to recover main window")
                    return False, last_y_position
            else:
                logging.error("No Chrome windows found for recovery")
                return False, last_y_position
        
        # Track button positions using our improved mechanism
        new_y, success = click_website_button(previous_y=last_y_position, button_positions=button_positions, button_index=button_index)
        if not success:
            return False, last_y_position
        
        # Wait longer for the new tab to open completely
        time.sleep(4)
        
        # Try different tab switching methods
        tab_switch_success = False
        
        # Method 1: Use ctrl+tab
        try:
            pyautogui.hotkey('ctrl', 'tab')
            time.sleep(2)
            current_window = gw.getActiveWindow()
            if current_window and original_title not in current_window.title:
                tab_switch_success = True
                logging.info("Successfully switched to new tab using ctrl+tab")
        except Exception as tab_error:
            logging.warning(f"Error switching tabs with ctrl+tab: {str(tab_error)}")
        
        # Method 2: If first method failed, try cycling through tabs
        if not tab_switch_success:
            try:
                # Try to find the newly opened tab by cycling through tabs
                for _ in range(5):  # Try up to 5 tabs
                    pyautogui.hotkey('ctrl', 'tab')
                    time.sleep(1)
                    current_window = gw.getActiveWindow()
                    if current_window and original_title not in current_window.title:
                        tab_switch_success = True
                        logging.info("Successfully switched to new tab by cycling")
                        break
            except Exception as cycle_error:
                logging.warning(f"Error cycling through tabs: {str(cycle_error)}")
        
        try:
            # Get the current URL with our improved method
            current_url = get_current_url()
            
            # If URL retrieval failed, try again after a short delay
            if not current_url:
                time.sleep(2)
                current_url = get_current_url()
                if not current_url:
                    logging.error("Failed to retrieve URL after multiple attempts")
                    # Close tab and return to main tab
                    pyautogui.hotkey('ctrl', 'w')
                    time.sleep(1)
                    return_to_main_tab(original_title)
                    return False, new_y
            
            base_url = get_base_url(current_url)
            
            if current_url in visited_urls or base_url in visited_urls:
                logging.info(f"Skipping duplicate: {current_url}")
                # Close tab and return
                pyautogui.hotkey('ctrl', 'w')
                time.sleep(1.5)  # Increased delay
                
                # Return to main results tab using tab cycling and title matching
                return_to_main_tab(original_title)
                
                return False, new_y
            
            # Add both full URL and base URL to visited_urls
            visited_urls.add(current_url)
            visited_urls.add(base_url)
            
            # Save to visited_urls.txt immediately to prevent duplicates across runs
            try:
                with open(visited_urls_file, 'a') as file:
                    file.write(current_url + '\n')
                    file.flush()  # Force write to disk immediately
                    os.fsync(file.fileno())  # Ensure it's written to the physical disk
            except Exception as file_error:
                logging.error(f"Error saving to visited_urls.txt: {str(file_error)}")
            
            page_title = get_page_title()
            
            # Perform scraping with updated return values
            has_chatbot, phone, email, contact_url, address, description, location = visit_and_check_website(current_url, page_title)
            
            # Save to CSV with immediate flush to ensure data is written
            try:
                save_to_csv(search_term, page_title, current_url, has_chatbot, phone, email, contact_url, address, description, location)
            except Exception as csv_error:
                logging.error(f"Error saving to CSV: {str(csv_error)}")
            
        except Exception as e:
            logging.error(f"Processing error: {str(e)}")
        finally:
            # Make sure we're on the correct tab before closing
            try:
                # Get all Chrome windows to check if we're on the right tab
                chrome_windows = gw.getWindowsWithTitle("Google Chrome")
                current_window = gw.getActiveWindow()
                
                # Only close tab if we're not on the main search tab
                if current_window and original_title not in current_window.title:
                    # Close tab and return
                    pyautogui.hotkey('ctrl', 'w')
                    time.sleep(1.5)  # Increased delay
                
                # Return to main results tab using tab cycling and title matching
                return_to_main_tab(original_title)
            except Exception as tab_error:
                logging.error(f"Error managing tabs: {str(tab_error)}")
                # Try to recover by finding a Chrome window with Google in the title
                try:
                    chrome_windows = gw.getWindowsWithTitle("Google Chrome")
                    for window in chrome_windows:
                        if "Google" in window.title:
                            window.activate()
                            break
                except Exception:
                    pass
            
            return True, new_y
    
    except Exception as e:
        logging.error(f"Process error: {str(e)}")
        return False, last_y_position


def return_to_main_tab(original_title):
    """Attempt to return to the main search results tab using multiple strategies"""
    if not original_title:
        logging.warning("No original title available for tab navigation")
        return False
    
    # First strategy: Try to find a window with the original title
    max_attempts = 10
    for attempt in range(max_attempts):
        try:
            # Get all Chrome windows
            chrome_windows = gw.getWindowsWithTitle("Google Chrome")
            if not chrome_windows:
                logging.warning("No Chrome windows found, trying to recover")
                time.sleep(1)
                continue
                
            current_window = gw.getActiveWindow()
            
            if current_window and original_title in current_window.title:
                logging.info("Already on the correct tab")
                return True
            
            # Check if any window has our original title
            for window in chrome_windows:
                if original_title in window.title:
                    window.activate()
                    time.sleep(1)  # Increased delay for reliable activation
                    logging.info(f"Found and activated original tab with title: {window.title}")
                    return True
            
            # If we didn't find the exact title, try cycling through tabs
            logging.info(f"Original tab not found by title, cycling through tabs (attempt {attempt+1}/{max_attempts})")
            
            # Try to identify the tab by looking for 'Google Search' or 'plumbers new york' in the title
            for window in chrome_windows:
                if ("Google Search" in window.title or 
                    "plumbers new york" in window.title.lower() or
                    "Google Maps" in window.title):
                    window.activate()
                    time.sleep(1)
                    logging.info(f"Found Google search tab by title keywords: {window.title}")
                    return True
            
            # If still not found, cycle through tabs
            pyautogui.hotkey('ctrl', 'tab')
            time.sleep(1)  # Increased delay
            
            # Check if we're now on a Google search results page
            current_window = gw.getActiveWindow()
            if current_window and ("Google Search" in current_window.title or 
                                  "Google Maps" in current_window.title or
                                  "plumbers new york" in current_window.title.lower()):
                logging.info(f"Found Google search tab: {current_window.title}")
                return True
            
        except Exception as e:
            logging.warning(f"Error during tab navigation attempt {attempt+1}: {str(e)}")
        
        time.sleep(1)  # Increased delay between attempts
    
    # Second strategy: Try to use keyboard shortcuts to get back to the first tab
    logging.warning("Using second strategy for tab navigation")
    try:
        # First go to the first tab
        pyautogui.hotkey('ctrl', '1')
        time.sleep(1.5)  # Increased delay
        
        # Check if we're on a Google search page
        current_window = gw.getActiveWindow()
        if current_window and ("Google Search" in current_window.title or 
                              "Google Maps" in current_window.title or
                              "plumbers new york" in current_window.title.lower()):
            logging.info(f"Found Google search tab using ctrl+1: {current_window.title}")
            return True
        
        # If not, try cycling through tabs
        for tab_num in range(2, 6):  # Try tabs 2-5
            pyautogui.hotkey('ctrl', str(tab_num))
            time.sleep(1.5)
            current_window = gw.getActiveWindow()
            if current_window and ("Google Search" in current_window.title or 
                                  "Google Maps" in current_window.title or
                                  "plumbers new york" in current_window.title.lower()):
                logging.info(f"Found Google search tab using ctrl+{tab_num}: {current_window.title}")
                return True
    except Exception as e:
        logging.error(f"Second strategy tab navigation failed: {str(e)}")
    
    # Last resort: Try to use alt+tab to switch between applications
    logging.warning("Using last resort tab navigation with alt+tab")
    try:
        # Try alt+tab to cycle through windows
        pyautogui.hotkey('alt', 'tab')
        time.sleep(1.5)
        
        # Check if we're on a Google window
        current_window = gw.getActiveWindow()
        if current_window and "Google Chrome" in current_window.title:
            # Now try to find the right tab
            for _ in range(10):  # Try up to 10 tabs
                pyautogui.hotkey('ctrl', 'tab')
                time.sleep(1)
                current_window = gw.getActiveWindow()
                if current_window and ("Google Search" in current_window.title or 
                                      "Google Maps" in current_window.title or
                                      "plumbers new york" in current_window.title.lower()):
                    logging.info(f"Found Google search tab after alt+tab: {current_window.title}")
                    return True
    except Exception as e:
        logging.error(f"Last resort tab navigation failed: {str(e)}")
    
    logging.error("Failed to return to main tab after multiple attempts")
    return False

def process_search_term(search_term):
    """Process a single search term"""
    logging.info(f"\n{'#'*50}\nStarting search for: {search_term}\n{'#'*50}")
    
    visited_urls = load_visited_urls(visited_urls_file)
    main_window = open_search_engine(search_term)
    
    if not main_window:
        return
    
    # Store the original window title for tab tracking
    original_title = None
    try:
        original_title = main_window.title
        logging.info(f"Original search window title: {original_title}")
    except Exception as e:
        logging.warning(f"Could not get original window title: {str(e)}")
    
    # Click on Places tab
    if not click_places_tab():
        logging.error(f"Could not access Places tab for '{search_term}'")
        return
    
    # Update the original title after clicking Places tab as it may have changed
    try:
        main_window = gw.getActiveWindow()
        original_title = main_window.title
        logging.info(f"Updated search window title after Places tab: {original_title}")
    except Exception as e:
        logging.warning(f"Could not update window title: {str(e)}")
    
    for page in range(1, max_pages + 1):
        logging.info(f"Processing Places page {page} for '{search_term}'")
        
        # First, count all website buttons on the page
        try:
            main_window.activate()
        except Exception as window_error:
            logging.error(f"Window activation error: {str(window_error)}")
            # Try to recover by finding a Chrome window
            chrome_windows = gw.getWindowsWithTitle("Google Chrome")
            if chrome_windows:
                main_window = chrome_windows[0]
                try:
                    main_window.activate()
                except Exception:
                    logging.error("Failed to recover main window")
                    continue  # Skip to next page
            else:
                logging.error("No Chrome windows found for recovery")
                break  # Exit the loop if no recovery possible
        
        time.sleep(1)
        
        # Scroll to the top of results
        for _ in range(5):  # Ensure we're at the top
            pyautogui.press('home')
            time.sleep(0.5)
        
        # Count all website buttons on the current page with progressive scrolling
        logging.info(f"Scanning for website buttons with progressive scrolling")
        all_buttons = count_website_buttons()
        
        num_buttons = len(all_buttons)
        logging.info(f"Found {num_buttons} website buttons on page {page}")
        
        if num_buttons == 0:
            logging.warning(f"No website buttons found on page {page}, trying fallback approach")
            # Fallback to old approach if no buttons found
            processed_count = 0
            max_businesses_per_page = 10
            last_y_position = None
            
            with alive_bar(max_businesses_per_page, title=f'Places Page {page} (Fallback)', bar='filling', spinner='dots_waves') as bar:
                for business_idx in range(max_businesses_per_page):
                    # Ensure we're on the main results tab
                    if original_title:
                        return_to_main_tab(original_title)
                    
                    try:
                        main_window = gw.getActiveWindow()
                        main_window.activate()
                    except Exception as window_error:
                        logging.warning(f"Window activation error in fallback: {str(window_error)}")
                        # Try to recover by finding a Chrome window
                        chrome_windows = gw.getWindowsWithTitle("Google Chrome")
                        if chrome_windows:
                            main_window = chrome_windows[0]
                            try:
                                main_window.activate()
                            except Exception:
                                logging.error("Failed to recover main window in fallback")
                                continue  # Skip this business
                        else:
                            logging.error("No Chrome windows found for recovery in fallback")
                            break  # Exit the loop if no recovery possible
                    
                    time.sleep(1)
                    
                    try:
                        result, last_y_position = process_business_listing(main_window, visited_urls, search_term, last_y_position)
                        if result:
                            processed_count += 1
                    except Exception as e:
                        logging.error(f"Error processing business in fallback: {str(e)}")
                        # Continue with next business
                    
                    bar()
                    
                    # Scroll down after every few businesses
                    if business_idx % 3 == 2:
                        # Ensure we're on the main results tab
                        if original_title:
                            return_to_main_tab(original_title)
                        
                        try:
                            main_window = gw.getActiveWindow()
                            main_window.activate()
                            time.sleep(1)
                            scroll_down_for_more_results()
                        except Exception as scroll_error:
                            logging.warning(f"Error scrolling in fallback: {str(scroll_error)}")
                    
                    time.sleep(2)
        else:
            # Process each button systematically
            with alive_bar(num_buttons, title=f'Places Page {page}', bar='filling', spinner='dots_waves') as bar:
                processed_count = 0
                
                # Scroll back to top before starting
                for _ in range(5):
                    pyautogui.press('home')
                    time.sleep(0.5)
                
                # Process each button one by one
                for button_idx in range(num_buttons):
                    # Ensure we're on the main results tab
                    if original_title:
                        return_to_main_tab(original_title)
                    
                    try:
                        main_window = gw.getActiveWindow()
                        main_window.activate()
                    except Exception as window_error:
                        logging.warning(f"Window activation error for button {button_idx}: {str(window_error)}")
                        # Try to recover by finding a Chrome window
                        chrome_windows = gw.getWindowsWithTitle("Google Chrome")
                        if chrome_windows:
                            main_window = chrome_windows[0]
                            try:
                                main_window.activate()
                            except Exception:
                                logging.error(f"Failed to recover main window for button {button_idx}")
                                bar()  # Update progress bar
                                continue  # Skip this button
                        else:
                            logging.error("No Chrome windows found for recovery")
                            break  # Exit the loop if no recovery possible
                    
                    time.sleep(1)
                    
                    try:
                        # Process current business with our improved tracking
                        result, _ = process_business_listing(
                            main_window, 
                            visited_urls, 
                            search_term, 
                            None,  # No need for last_y_position when using pre-counted buttons
                            all_buttons, 
                            button_idx
                        )
                        
                        if result:
                            processed_count += 1
                    except Exception as e:
                        logging.error(f"Error processing button {button_idx}: {str(e)}")
                        # Continue with next button
                    
                    bar()
                    
                    # Scroll down periodically to ensure buttons remain visible
                    if button_idx > 0 and button_idx % 3 == 0:
                        # Ensure we're on the main results tab
                        if original_title:
                            return_to_main_tab(original_title)
                        
                        try:
                            main_window = gw.getActiveWindow()
                            main_window.activate()
                            time.sleep(1)
                            scroll_down_for_more_results()
                            time.sleep(2)
                            
                            # Re-count buttons after scrolling to ensure we have all
                            new_buttons = count_website_buttons()
                            if len(new_buttons) > len(all_buttons):
                                # Add any new buttons we found
                                existing_positions = {(x, y) for x, y in all_buttons}
                                for x, y in new_buttons:
                                    if (x, y) not in existing_positions:
                                        all_buttons.append((x, y))
                                num_buttons = len(all_buttons)
                                logging.info(f"Updated button count to {num_buttons}")
                        except Exception as scroll_error:
                            logging.warning(f"Error during scroll or button recount: {str(scroll_error)}")
        
        logging.info(f"Processed {processed_count} businesses on page {page}")
        
        # Save visited URLs to file
        save_visited_urls(visited_urls_file, visited_urls)
        
        # Navigate to the next page if not on the last page
        if page < max_pages:
            # Ensure we're on the main results tab
            if original_title:
                return_to_main_tab(original_title)
            
            # Make sure main window is active
            try:
                main_window = gw.getActiveWindow()
                main_window.activate()
                time.sleep(1)
                if not navigate_to_next_page():
                    logging.warning(f"Could not navigate to page {page+1}, stopping")
                    break
            except Exception as nav_error:
                logging.error(f"Error navigating to next page: {str(nav_error)}")
                break
    
    # Close the main window
    try:
        main_window.close()
    except Exception as close_error:
        logging.warning(f"Error closing main window: {str(close_error)}")
        # Try to close any remaining Chrome windows
        try:
            chrome_windows = gw.getWindowsWithTitle("Google Chrome")
            for window in chrome_windows:
                window.close()
        except:
            pass

def get_current_url():
    """Get the URL of the current tab using clipboard"""
    # Use keyboard shortcut to copy current URL to clipboard
    pyautogui.hotkey('alt', 'd')  # Select address bar
    time.sleep(0.8)  # Increased delay for reliability
    pyautogui.hotkey('ctrl', 'a')  # Select all text in address bar
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c')  # Copy URL to clipboard
    time.sleep(0.8)  # Increased delay for reliability
    
    # Get URL from clipboard with retry mechanism
    max_attempts = 3
    for attempt in range(max_attempts):
        url = pyperclip.paste()
        if url and url.startswith('http'):
            logging.info(f"Retrieved URL from clipboard: {url}")
            return url
        else:
            logging.warning(f"Failed to get valid URL on attempt {attempt+1}, retrying...")
            # Try again with different approach
            pyautogui.hotkey('alt', 'd')  # Select address bar
            time.sleep(0.8)
            pyautogui.hotkey('ctrl', 'c')  # Copy URL to clipboard
            time.sleep(0.8)
    
    # If we get here, we couldn't get a valid URL
    logging.error("Failed to retrieve valid URL after multiple attempts")
    return ""  # Return empty string instead of None to avoid errors

def get_page_title():
    """Get the title of the current page"""
    try:
        current_window = gw.getActiveWindow()
        if current_window:
            title = current_window.title
            # Remove browser name from title if present
            if ' - Google Chrome' in title:
                title = title.replace(' - Google Chrome', '')
            return title
    except Exception as e:
        logging.error(f"Error getting page title: {str(e)}")
    
    return "Unknown Title"

def scroll_down_for_more_results():
    """Scroll down to reveal more results"""
    try:
        # Scroll down gradually in smaller increments for smoother scrolling
        for _ in range(3):
            pyautogui.scroll(-250)  # Smaller increments
            time.sleep(0.5)  # Short pause between scrolls
        
        # Wait for page to settle and load new content
        time.sleep(1.5)  # Increased wait time
        
        logging.info("Scrolled down for more results")
        return True
    except Exception as e:
        logging.error(f"Error scrolling for more results: {str(e)}")
        # Try alternative approach if standard scroll fails
        try:
            # Use Page Down key as alternative
            pyautogui.press('pagedown')
            time.sleep(1.5)
            logging.info("Used Page Down key as alternative scroll method")
            return True
        except Exception as alt_error:
            logging.error(f"Alternative scroll method failed: {str(alt_error)}")
            return False

def scroll_down_for_more_results():
    """Scroll down to reveal more results"""
    try:
        # Scroll down gradually in smaller increments for smoother scrolling
        for _ in range(3):
            pyautogui.scroll(-250)  # Smaller increments
            time.sleep(0.5)  # Short pause between scrolls
        
        # Wait for page to settle and load new content
        time.sleep(1.5)  # Increased wait time
        
        logging.info("Scrolled down for more results")
        return True
    except Exception as e:
        logging.error(f"Error scrolling for more results: {str(e)}")
        # Try alternative approach if standard scroll fails
        try:
            # Use Page Down key as alternative
            pyautogui.press('pagedown')
            time.sleep(1.5)
            logging.info("Used Page Down key as alternative scroll method")
            return True
        except Exception as alt_error:
            logging.error(f"Alternative scroll method failed: {str(alt_error)}")
            return False

def main():
    # Initialize the CSV file
    initialize_csv()
    
    # Process each search term with error handling
    for search_term in search_terms:
        try:
            logging.info(f"Starting to process search term: {search_term}")
            process_search_term(search_term)
            time.sleep(5)  # Pause between search terms
        except Exception as e:
            logging.error(f"Error processing search term '{search_term}': {str(e)}")
            # Try to close any hanging Chrome windows
            try:
                chrome_windows = gw.getWindowsWithTitle("Google Chrome")
                for window in chrome_windows:
                    try:
                        window.close()
                    except:
                        pass
            except Exception as window_error:
                logging.error(f"Error closing Chrome windows: {str(window_error)}")
            
            # Continue with next search term
            time.sleep(10)  # Longer pause after error
            continue
    
    logging.info("Script completed successfully")

if __name__ == "__main__":
    main()

def navigate_to_next_page():
    """Navigate to the next page of search results"""
    try:
        logging.info("Navigating to next page")
        
        # First, scroll to the bottom of the page to ensure next button is visible
        for _ in range(3):
            pyautogui.scroll(-500)  # Scroll down in smaller increments
            time.sleep(0.8)
        
        # Try to find the next page button with multiple confidence levels
        for confidence in [0.8, 0.7, 0.6, 0.5]:
            next_button = locate_button("img/next_page.png", timeout=3, confidence=confidence)
            if next_button:
                logging.info(f"Found next page button using img/next_page.png with confidence {confidence}")
                center_x, center_y = pyautogui.center(next_button)
                # Move mouse to button position first
                pyautogui.moveTo(center_x, center_y)
                time.sleep(0.5)  # Small pause before clicking
                pyautogui.click()
                logging.info(f"Clicked next page button at position ({center_x}, {center_y})")
                time.sleep(6)  # Increased wait time for page to load
                return True
        
        # Try alternative image if first one fails
        for confidence in [0.7, 0.6, 0.5]:
            next_button = locate_button("img/next_pagev.png", timeout=3, confidence=confidence)
            if next_button:
                logging.info(f"Found next page button using img/next_pagev.png with confidence {confidence}")
                center_x, center_y = pyautogui.center(next_button)
                # Move mouse to button position first
                pyautogui.moveTo(center_x, center_y)
                time.sleep(0.5)  # Small pause before clicking
                pyautogui.click()
                logging.info(f"Clicked next page button at position ({center_x}, {center_y})")
                time.sleep(6)  # Increased wait time for page to load
                return True
        
        # Try looking for the button in the bottom right area of the screen
        screen_width, screen_height = pyautogui.size()
        bottom_right_region = (screen_width // 2, screen_height - 300, screen_width // 2, 300)
        
        for confidence in [0.6, 0.5, 0.4]:
            next_button = locate_button("img/next_page.png", timeout=3, confidence=confidence, region=bottom_right_region)
            if next_button:
                logging.info(f"Found next page button in bottom right region with confidence {confidence}")
                center_x, center_y = pyautogui.center(next_button)
                pyautogui.click(center_x, center_y)
                logging.info(f"Clicked next page button at position ({center_x}, {center_y})")
                time.sleep(6)  # Increased wait time for page to load
                return True
        
        # Try fallback approach - look for the '>' symbol at bottom of page
        logging.info("Trying fallback approach for next page")
        # Click in the area where the next button typically is in Google Maps results
        target_x = screen_width // 2 + 250
        target_y = screen_height - 150
        pyautogui.moveTo(target_x, target_y)
        time.sleep(0.5)
        pyautogui.click()
        logging.info(f"Used fallback click at position ({target_x}, {target_y})")
        time.sleep(6)  # Increased wait time
        return True
    except Exception as e:
        logging.error(f"Error navigating to next page: {str(e)}")
        return False

def visit_and_check_website(url, page_title):
    """Visit website and check for chatbot, contact info, etc."""
    has_chatbot = False
    phone = ""
    email = ""
    contact_url = ""
    address = ""
    description = ""
    location = ""
    
    try:
        # Try to get the page content
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Check for chatbot markers in HTML
            page_html = response.text.lower()
            for marker in chatbot_markers:
                if marker in page_html:
                    has_chatbot = True
                    logging.info(f"Chatbot detected on {url} with marker: {marker}")
                    break
            
            # Extract meta description if available
            meta_desc = soup.find('meta', attrs={'name': 'description'})
            if meta_desc and 'content' in meta_desc.attrs:
                description = meta_desc['content']
            
            # Look for contact information
            # Phone numbers - look for common patterns
            phone_patterns = [
                r'\(\d{3}\)\s*\d{3}-\d{4}',  # (123) 456-7890
                r'\d{3}-\d{3}-\d{4}',          # 123-456-7890
                r'\d{3}\.\d{3}\.\d{4}',        # 123.456.7890
                r'\+\d{1,2}\s*\(\d{3}\)\s*\d{3}-\d{4}',  # +1 (123) 456-7890
                r'\+\d{1,2}\s*\d{3}\s*\d{3}\s*\d{4}'     # +1 123 456 7890
            ]
            
            for pattern in phone_patterns:
                import re
                phone_matches = re.findall(pattern, response.text)
                if phone_matches:
                    phone = phone_matches[0]
                    break
            
            # Email addresses
            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            email_matches = re.findall(email_pattern, response.text)
            if email_matches:
                # Filter out common false positives
                filtered_emails = [e for e in email_matches if not (
                    'example' in e or 
                    'your@email' in e or 
                    'user@' in e or 
                    'name@' in e
                )]
                if filtered_emails:
                    email = filtered_emails[0]
            
            # Look for contact page link
            contact_links = soup.find_all('a', href=True, text=lambda text: text and 'contact' in text.lower())
            if contact_links:
                contact_path = contact_links[0]['href']
                # Convert relative URL to absolute if needed
                contact_url = urljoin(url, contact_path)
            
            # Try to find location/address information
            address_candidates = soup.find_all(['p', 'div', 'span'], text=lambda text: text and ('address' in text.lower() or 'location' in text.lower()))
            if address_candidates:
                for candidate in address_candidates:
                    if len(candidate.text) > 10 and len(candidate.text) < 200:  # Reasonable address length
                        address = candidate.text.strip()
                        break
            
            # Look for location information
            location_elements = soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'div'], 
                                           text=lambda text: text and ('location' in text.lower() or 'our office' in text.lower()))
            if location_elements:
                for elem in location_elements:
                    if 'location' in elem.text.lower() or 'our office' in elem.text.lower():
                        location = elem.text.strip()
                        break
        else:
            logging.warning(f"Failed to retrieve {url}: Status code {response.status_code}")
    except Exception as e:
        logging.error(f"Error visiting {url} for {page_title}: {str(e)}")
    
    return has_chatbot, phone, email, contact_url, address, description, location
