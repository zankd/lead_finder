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
        with open(file_path, 'r') as file:
            for line in file:
                url = line.strip()
                if url:
                    existing_urls.add(url)
    
    # Combine with new URLs
    all_urls = existing_urls.union(visited_urls)
    
    # Write all URLs back to file
    with open(file_path, 'w') as file:
        for url in all_urls:
            file.write(url + '\n')
    
    logging.info(f"Saved {len(all_urls)} URLs to {file_path}")

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
    
    with open(results_csv, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # Ensure the order matches the desired format: Search Term,URL,Title,Description,Has Chatbot,Phone,Email,Contact URL,Location
        writer.writerow([search_term, url, business_name, description, has_chatbot_formatted, formatted_phone, email, contact_url, cleaned_location])
    logging.info(f"Saved data for {business_name} - {url}")

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
    """Count all website buttons visible on the current page"""
    screen_width, screen_height = pyautogui.size()
    all_buttons = []
    y_position = 300  # Start from top of results
    
    # Scan the page in sections to find all buttons
    while y_position < screen_height - 100:
        region = (0, y_position, screen_width, 300)  # Search in a window of 300px height
        try:
            button = pyautogui.locateOnScreen("img/webss.png", confidence=0.7, region=region)
            if button:
                center_x, center_y = pyautogui.center(button)
                # Check if this button is already in our list (avoid duplicates)
                if not any(abs(center_y - y) < 20 for _, y in all_buttons):
                    all_buttons.append((center_x, center_y))
                y_position = center_y + 50  # Move past this button
            else:
                y_position += 200  # Move down if no button found
        except Exception as e:
            logging.warning(f"Error during button counting: {str(e)}")
            y_position += 200  # Move down on error
    
    logging.info(f"Found {len(all_buttons)} website buttons on current page")
    return all_buttons

def click_website_button(timeout=5, previous_y=None, button_positions=None, button_index=0):
    """Click website button with position tracking"""
    screen_width, screen_height = pyautogui.size()
    
    # If we have pre-counted button positions, use those
    if button_positions and button_index < len(button_positions):
        center_x, center_y = button_positions[button_index]
        logging.info(f"Clicking pre-counted button {button_index+1}/{len(button_positions)} at position ({center_x}, {center_y})")
        
        # Open in new tab
        pyautogui.moveTo(center_x, center_y)
        pyautogui.keyDown('ctrl')
        pyautogui.click()
        pyautogui.keyUp('ctrl')
        
        # Move cursor to avoid hover effects
        pyautogui.moveTo(center_x, center_y + result_offset)
        time.sleep(3)
        return center_y + result_offset, True  # Return position and success flag
    
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
        center_x, center_y = pyautogui.center(website_button)
        
        # Open in new tab
        pyautogui.moveTo(center_x, center_y)
        pyautogui.keyDown('ctrl')
        pyautogui.click()
        pyautogui.keyUp('ctrl')
        
        # Move cursor to avoid hover effects
        pyautogui.moveTo(center_x, center_y + result_offset)
        time.sleep(3)
        return center_y + result_offset, True  # Return last clicked position and success flag
    else:
        logging.warning("Button not found: img/webss.png")
        return None, False

def process_business_listing(main_window, visited_urls, search_term, last_y_position, button_positions=None, button_index=0):
    try:
        main_window.activate()
        time.sleep(1)
        
        # Track button positions using our improved mechanism
        new_y, success = click_website_button(previous_y=last_y_position, button_positions=button_positions, button_index=button_index)
        if not success:
            return False, last_y_position
        
        time.sleep(3)
        
        # Switch to new tab
        pyautogui.hotkey('ctrl', 'tab')
        time.sleep(1)
        
        try:
            current_url = get_current_url()
            base_url = get_base_url(current_url)
            
            if current_url in visited_urls or base_url in visited_urls:
                logging.info(f"Skipping duplicate: {current_url}")
                # Close tab and return
                pyautogui.hotkey('ctrl', 'w')
                time.sleep(1)
                main_window.activate()
                return False, new_y
            
            # Add both full URL and base URL to visited_urls
            visited_urls.add(current_url)
            visited_urls.add(base_url)
            
            # Save to visited_urls.txt immediately to prevent duplicates across runs
            with open(visited_urls_file, 'a') as file:
                file.write(current_url + '\n')
            
            page_title = get_page_title()
            
            # Perform scraping with updated return values
            has_chatbot, phone, email, contact_url, address, description, location = visit_and_check_website(current_url, page_title)
            save_to_csv(search_term, page_title, current_url, has_chatbot, phone, email, contact_url, address, description, location)
            
        except Exception as e:
            logging.error(f"Processing error: {str(e)}")
        finally:
            # Close tab and return
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(1)
            main_window.activate()
            
            return True, new_y
    
    except Exception as e:
        logging.error(f"Process error: {str(e)}")
        return False, last_y_position

def process_search_term(search_term):
    """Process a single search term"""
    logging.info(f"\n{'#'*50}\nStarting search for: {search_term}\n{'#'*50}")
    
    visited_urls = load_visited_urls(visited_urls_file)
    main_window = open_search_engine(search_term)
    
    if not main_window:
        return
    
    # Click on Places tab
    if not click_places_tab():
        logging.error(f"Could not access Places tab for '{search_term}'")
        return
    
    for page in range(1, max_pages + 1):
        logging.info(f"Processing Places page {page} for '{search_term}'")
        
        # Process multiple business listings on the current page
        processed_count = 0
        max_businesses_per_page = 10
        
        with alive_bar(max_businesses_per_page, title=f'Places Page {page}', bar='filling', spinner='dots_waves') as bar:
            last_y = None  # Initialize position tracking
            for business_idx in range(max_businesses_per_page):
                # Make sure main window is active
                main_window.activate()
                time.sleep(1)
                
                # Process current business with position tracking
                result, new_y = process_business_listing(main_window, visited_urls, search_term, last_y)
                if result:
                    processed_count += 1
                    last_y = new_y  # Update position tracking
                
                bar()
                
                if business_idx % 2 == 1:
                    main_window.activate()
                    time.sleep(1)
                    scroll_down_for_more_results()
                    time.sleep(2)
                    last_y = None
                
                time.sleep(1.5)
        
        logging.info(f"Processed {processed_count} businesses on page {page}")
        
        # Navigate to the next page if not on the last page
        if page < max_pages:
            # Make sure main window is active
            main_window.activate()
            time.sleep(1)
            if not navigate_to_next_page():
                logging.warning(f"Could not navigate to page {page+1}, stopping")
                break
    
    # Close the main window
    try:
        main_window.close()
    except:
        pass

def get_current_url():
    pyautogui.hotkey('ctrl', 'l')
    time.sleep(0.8)
    
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.7)
    
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.6)
    
    url = pyperclip.paste().strip()
    
    if not url.startswith(('http://', 'https://')):
        logging.warning(f"Invalid URL format: {url}")
        return ""
    
    return url

def get_page_title():
    """Get the page title"""
    active_window = gw.getActiveWindow()
    if not active_window:
        return "Unknown"
    
    title = active_window.title
    if " - " in title:
        title = title.split(" - ")[0]
    return title

def detect_chatbot(soup):
    for marker in chatbot_markers:
        elements = soup.select(f'[id*="{marker}"], [class*="{marker}"], [data-*="{marker}"]')
        if elements:
            return True
    
    iframes = soup.find_all('iframe')
    for iframe in iframes:
        src = iframe.get('src', '')
        if any(bot_provider in src for bot_provider in ['chat-widget', 'livechat', 'chat-bot', 'chatbot', 'live-chat', 
    'bot-widget', 'zendesk', 'intercom', 'drift', 'freshchat',
    'crisp', 'tawk', 'liveperson', 'olark', 'userlike',
    'chat_widget', 'chat_container', 'chat_bubble', 'chat_frame']):
            return True
    
    scripts = soup.find_all('script')
    for script in scripts:
        src = script.get('src', '')
        if any(bot_provider in src for bot_provider in ['chat-widget', 'livechat', 'chat-bot', 'chatbot', 'live-chat', 
    'bot-widget', 'zendesk', 'intercom', 'drift', 'freshchat',
    'crisp', 'tawk', 'liveperson', 'olark', 'userlike',
    'chat_widget', 'chat_container', 'chat_bubble', 'chat_frame']):
            return True
    
    return False

def extract_contact_info(soup, url):
    phone = ""
    email = ""
    contact_url = ""
    address = ""
    
    phone_elements = soup.select('a[href^="tel:"]')
    if phone_elements:
        phone = phone_elements[0].get('href').replace('tel:', '')
    
    # Extract email addresses
    email_elements = soup.select('a[href^="mailto:"]')
    if email_elements:
        email = email_elements[0].get('href').replace('mailto:', '')
    
    # Look for contact page link
    contact_links = soup.find_all('a', string=lambda s: s and ('contact' in s.lower() or 'get in touch' in s.lower()))
    if contact_links:
        contact_href = contact_links[0].get('href', '')
        if contact_href:
            contact_url = urljoin(url, contact_href)
    
    # Look for address information
    address_elements = soup.find_all(['p', 'div', 'span'], 
                                    string=lambda s: s and any(x in s.lower() for x in ['address', 'location', 'street', 'avenue', 'blvd', 'road']))
    if address_elements:
        address = address_elements[0].get_text().strip()
    
    return phone, email, contact_url, address

def visit_and_check_website(url, business_name):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=15)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        has_chatbot = detect_chatbot(soup)
        
        # Extract contact information
        phone, email, contact_url, address = extract_contact_info(soup, url)
        
        # Extract description
        description = ""
        meta_desc = soup.find('meta', {'name': 'description'}) or soup.find('meta', {'property': 'og:description'})
        if meta_desc:
            description = meta_desc.get('content', '').strip()
        if not description:
            # Try to get first paragraph or div with substantial text
            for tag in soup.find_all(['p', 'div']):
                text = tag.get_text().strip()
                if len(text) > 100:  # Only consider substantial text
                    description = text
                    break
        
        # Extract location from address if not explicitly provided
        location = address
        if not location:
            location_elements = soup.find_all(['p', 'div', 'span'],
                string=lambda s: s and any(x in s.lower() for x in ['location', 'based in', 'located in', 'serving']))
            if location_elements:
                location = location_elements[0].get_text().strip()
        
        for _ in range(3):
            pyautogui.scroll(-500)
            time.sleep(1)
        
        return has_chatbot, phone, email, contact_url, address, description, location
    
    except Exception as e:
        logging.error(f"Error visiting {url} for {business_name}: {str(e)}")
        return False, "", "", "", "", "", ""

def scroll_down_for_more_results():
    try:
        screen_width, screen_height = pyautogui.size()
        
        pyautogui.press('pagedown', presses=2, interval=0.5)
        time.sleep(1)
        
        pyautogui.moveTo(screen_width//2, screen_height//2)
    except Exception as e:
        logging.error(f"Scrolling error: {str(e)}")

def navigate_to_next_page():
    logging.info("Navigating to next page")
    
    # Scroll down to find the next page button
    for _ in range(10):
        pyautogui.scroll(-700)
        time.sleep(0.5)
    
    # Try multiple next page button images with different confidence levels
    next_button = None
    for image_path in ["img/next_page.png", "img/next_pagev.png", "img/nn.png"]:
        if os.path.exists(image_path):
            for confidence in [0.8, 0.7, 0.6, 0.5]:
                next_button = locate_button(image_path, timeout=3, confidence=confidence)
                if next_button:
                    logging.info(f"Found next page button using {image_path} with confidence {confidence}")
                    break
        if next_button:
            break
    
    if next_button:
        # Click the next page button
        center_x, center_y = pyautogui.center(next_button)
        pyautogui.click(center_x, center_y)
        logging.info(f"Clicked next page button at position ({center_x}, {center_y})")
        time.sleep(5)  # Wait for page to load
        return True
    else:
        logging.warning("Next page button not found")
        return False

def process_search_term(search_term):
    """Process a single search term"""
    logging.info(f"\n{'#'*50}\nStarting search for: {search_term}\n{'#'*50}")
    
    visited_urls = load_visited_urls(visited_urls_file)
    main_window = open_search_engine(search_term)
    
    if not main_window:
        return
    
    # Click on Places tab
    if not click_places_tab():
        logging.error(f"Could not access Places tab for '{search_term}'")
        return
    
    for page in range(1, max_pages + 1):
        logging.info(f"Processing Places page {page} for '{search_term}'")
        
        # First, count all website buttons on the page
        main_window.activate()
        time.sleep(1)
        
        # Scroll to the top of results
        for _ in range(5):  # Ensure we're at the top
            pyautogui.press('home')
            time.sleep(0.5)
        
        # Count all website buttons on the current page
        all_buttons = []
        for scan_attempt in range(3):  # Multiple scan attempts to ensure we find all buttons
            logging.info(f"Scanning for website buttons (attempt {scan_attempt+1}/3)")
            
            # Count buttons
            buttons = count_website_buttons()
            if buttons:
                all_buttons = buttons
                break
            
            # If no buttons found, try scrolling and scanning again
            scroll_down_for_more_results()
            time.sleep(2)
        
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
                    main_window.activate()
                    time.sleep(1)
                    
                    result, last_y_position = process_business_listing(main_window, visited_urls, search_term, last_y_position)
                    if result:
                        processed_count += 1
                    
                    bar()
                    
                    # Scroll down after every few businesses
                    if business_idx % 3 == 2:
                        main_window.activate()
                        time.sleep(1)
                        scroll_down_for_more_results()
                    
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
                    main_window.activate()
                    time.sleep(1)
                    
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
                    
                    bar()
                    
                    # Scroll down periodically to ensure buttons remain visible
                    if button_idx > 0 and button_idx % 3 == 0:
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
        
        logging.info(f"Processed {processed_count} businesses on page {page}")
        
        # Save visited URLs to file
        save_visited_urls(visited_urls_file, visited_urls)
        
        # Navigate to the next page if not on the last page
        if page < max_pages:
            # Make sure main window is active
            main_window.activate()
            time.sleep(1)
            if not navigate_to_next_page():
                logging.warning(f"Could not navigate to page {page+1}, stopping")
                break
    
    # Close the main window
    try:
        main_window.close()
    except:
        pass

def main():
    # Initialize the CSV file
    initialize_csv()
    
    # Process each search term
    for search_term in search_terms:
        process_search_term(search_term)
        time.sleep(5)  # Pause between search terms
    
    logging.info("Script completed successfully")

if __name__ == "__main__":
    main()
