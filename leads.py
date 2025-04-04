import os, csv, json, time, logging
import pyautogui
import requests
import webbrowser
import pygetwindow as gw
import pyperclip
from bs4 import BeautifulSoup
from alive_progress import alive_bar
from urllib.parse import urljoin, urlparse
from hubspot_integration import send_data_to_hubspot

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraper_log_copy.txt'),
        logging.StreamHandler()
    ]
)

chrome_path = r"C:\Users\XnD\AppData\Local\Google\Chrome SxS\Application\chrome.exe"
visited_urls_file = 'visited_urls_copy.txt'
search_terms = [
    "plumbers new york",
    "electricians manhattan",
    "dentists london",
    "auto repair florida"
]
results_csv = 'website_results_copy.csv'
max_pages = 2
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
    if os.path.exists(file_path):
        with open(file_path, 'r') as file:
            return set(line.strip().split(',')[0] for line in file)
    return set()

def save_visited_urls(file_path, visited_urls):
    with open(file_path, 'w') as file:
        for url in visited_urls:
            file.write(url + '\n')

def initialize_csv():
    if not os.path.exists(results_csv):
        with open(results_csv, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Search Term', 'URL', 'Title', 'Description', 'Has Chatbot', 'Phone', 'Email', 'Contact URL', 'Location', 'Industry', 'City', 'Country'])

def save_to_csv(search_term, business_name, url, has_chatbot, phone, email, contact_url, address, description, location):
    has_chatbot_formatted = "YES" if has_chatbot else "NO"
    
    formatted_phone = phone
    if phone and not phone.startswith("+"):
        formatted_phone = "+1 " + phone
    
    cleaned_location = location
    if cleaned_location:
        if cleaned_location.lower().startswith("locations"):
            cleaned_location = cleaned_location[9:].strip()
        
        location_parts = cleaned_location.split(",")
        if len(location_parts) > 2:
            cleaned_location = ", ".join(location_parts[-2:]).strip()
        
        cleaned_location = cleaned_location[:100]
    
    # Extract industry and city from search term
    from country_mapping import extract_industry_and_city, get_country_from_city
    industry, city = extract_industry_and_city(search_term)
    country = get_country_from_city(city)
    
    # Save to CSV file
    with open(results_csv, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([search_term, url, business_name, description, has_chatbot_formatted, formatted_phone, email, contact_url, cleaned_location, industry, city, country])
    logging.info(f"Saved data for {business_name} - {url}")
    
    # Send data to HubSpot
    try:
        # Extract industry, city and country from search term
        from country_mapping import extract_industry_and_city, get_country_from_city
        industry, city = extract_industry_and_city(search_term)
        country = get_country_from_city(city)
        
        hubspot_result = send_data_to_hubspot(
            business_name=business_name,
            url=url,
            description=description,
            has_chatbot=has_chatbot,
            phone=formatted_phone,
            email=email,
            contact_url=contact_url,
            location=cleaned_location,
            industry=industry,
            city=city,
            country=country
        )
        if hubspot_result:
            logging.info(f"Successfully sent data to HubSpot for {business_name}")
        else:
            logging.warning(f"Failed to send data to HubSpot for {business_name}")
    except Exception as e:
        logging.error(f"Error sending data to HubSpot: {str(e)}")

def locate_button(image_path, timeout=10, region=None, confidence=0.8):
    start_time = time.time()
    screen_width, screen_height = pyautogui.size()
    
    original_image_path = image_path
    
    if not os.path.exists(image_path):
        logging.error(f"Image file not found: {image_path}")
        return None
    
    while time.time() - start_time < timeout:
        try:
            confidence_levels = [confidence, 0.7, 0.6, 0.5, 0.4, 0.3]
            for conf in confidence_levels:
                if region:
                    button = pyautogui.locateOnScreen(image_path, confidence=conf, region=region)
                    if button:
                        logging.info(f"Button found: {original_image_path} with confidence {conf}")
                        return button
                
                button = pyautogui.locateOnScreen(image_path, confidence=conf)
                if button:
                    logging.info(f"Button found: {original_image_path}")
                    return button
                
                time.sleep(0.2)
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as e:
            logging.warning(f"Image search error: {str(e)}")
        
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
    
    places_tab = locate_button("img/places.png", timeout=10, confidence=0.7)
    if places_tab:
        center_x, center_y = pyautogui.center(places_tab)
        logging.info(f"Places tab found at position: ({center_x}, {center_y})")
        pyautogui.click(places_tab)
        time.sleep(3)
        return True
    
    try:
        screen_width, screen_height = pyautogui.size()
        
        fallback_positions = [
            (screen_width // 4, 200),
            (480, 200),
            (screen_width // 4, 150),
            (screen_width // 4, 250)
        ]
        
        for pos_x, pos_y in fallback_positions:
            logging.info(f"Falling back to approximate position: ({pos_x}, {pos_y})")
            pyautogui.click(pos_x, pos_y)
            time.sleep(2)
            
            web_button = locate_button("img/webss.png", timeout=3, confidence=0.6)
            if web_button:
                logging.info("Places tab click successful (verified by finding web button)")
                return True
        
        logging.warning("All fallback positions failed, using last resort position")
        pyautogui.click(480, 200)
        time.sleep(3)
        return True
    except Exception as e:
        logging.error(f"Error clicking Places tab: {str(e)}")
        return False

scroll_step = 300
result_offset = 80

def click_website_button(timeout=5, previous_y=None):
    """Click website button with position tracking"""
    screen_width, screen_height = pyautogui.size()
    min_region_height = 200
    
    if previous_y is not None:
        region_height = max(screen_height - previous_y, min_region_height)
        region = (0, previous_y, screen_width, region_height)
    else:
        region = (0, 300, screen_width, screen_height - 300)

    website_button = None
    try:
        website_button = locate_button("img/webss.png", timeout=timeout, region=region)
    except Exception as e:
        logging.warning(f"Error locating button in region: {str(e)}")
    
    if not website_button:
        try:
            logging.info("Trying to find web button in full screen")
            website_button = locate_button("img/webss.png", timeout=timeout)
        except Exception as e:
            logging.warning(f"Error locating button in full screen: {str(e)}")
            return None

    if website_button:
        center_x, center_y = pyautogui.center(website_button)
        
        pyautogui.moveTo(center_x, center_y)
        pyautogui.keyDown('ctrl')
        pyautogui.click()
        pyautogui.keyUp('ctrl')
        
        pyautogui.moveTo(center_x, center_y + result_offset)
        time.sleep(3)
        return center_y + result_offset
    else:
        logging.warning("Button not found: img/webss.png")
        return None

def process_business_listing(main_window, visited_urls, search_term, last_y_position):
    try:
        main_window.activate()
        time.sleep(1)
        
        new_y = click_website_button(previous_y=last_y_position)
        if not new_y:
            return False, last_y_position
        
        time.sleep(3)
        
        pyautogui.hotkey('ctrl', 'tab')
        time.sleep(1)
        
        try:
            current_url = get_current_url()
            if current_url in visited_urls:
                logging.info(f"Skipping duplicate: {current_url}")
                return False, new_y
            
            visited_urls.add(current_url)
            page_title = get_page_title()
            
            company_name, has_chatbot, phone, email, contact_url, address, description, location = visit_and_check_website(current_url, page_title)
            save_to_csv(search_term, company_name, current_url, has_chatbot, phone, email, contact_url, address, description, location)
            
        except Exception as e:
            logging.error(f"Processing error: {str(e)}")
        finally:
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(1)
            main_window.activate()
            
            pyautogui.press('down')
            time.sleep(0.5)
            
            return True, new_y
    
    except Exception as e:
        logging.error(f"Process error: {str(e)}")
        return False, last_y_position

def process_search_term(search_term):
    logging.info(f"\n{'#'*50}\nStarting search for: {search_term}\n{'#'*50}")
    
    visited_urls = load_visited_urls(visited_urls_file)
    logging.info(f"Loaded {len(visited_urls)} visited URLs from {visited_urls_file}")
    main_window = open_search_engine(search_term)
    
    if not main_window:
        return
    
    if not click_places_tab():
        logging.error(f"Could not access Places tab for '{search_term}'")
        return
    
    for page in range(1, max_pages + 1):
        logging.info(f"Processing Places page {page} for '{search_term}'")
        
        processed_count = 0
        max_businesses_per_page = 10
        
        with alive_bar(max_businesses_per_page, title=f'Places Page {page}', bar='filling', spinner='dots_waves') as bar:
            last_y = None
            for business_idx in range(max_businesses_per_page):
                main_window.activate()
                time.sleep(1)
                
                result, new_y = process_business_listing(main_window, visited_urls, search_term, last_y)
                if result:
                    processed_count += 1
                    last_y = new_y
                
                bar()
                
                if business_idx % 2 == 1:
                    main_window.activate()
                    time.sleep(1)
                    scroll_down_for_more_results()
                    time.sleep(2)
                    last_y = None
                
                time.sleep(1.5)
        
        logging.info(f"Processed {processed_count} businesses on page {page}")
        
        save_visited_urls(visited_urls_file, visited_urls)
        logging.info(f"Saved {len(visited_urls)} URLs to {visited_urls_file}")
        
        if page < max_pages:
            main_window.activate()
            time.sleep(1)
            if not navigate_to_next_page():
                logging.warning(f"Could not navigate to page {page+1}, stopping")
                break
    
    try:
        current_tab = main_window.current_window_handle
        main_window.close_current_tab()
        remaining_handles = main_window.window_handles
        if remaining_handles:
            main_window.switch_to.window(remaining_handles[0])
    except Exception as e:
        logging.warning(f"Error while closing tab: {str(e)}")
        try:
            remaining_handles = main_window.window_handles
            if remaining_handles:
                main_window.switch_to.window(remaining_handles[0])
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
    
    email_elements = soup.select('a[href^="mailto:"]')
    if email_elements:
        email = email_elements[0].get('href').replace('mailto:', '')
    
    contact_links = soup.find_all('a', string=lambda s: s and ('contact' in s.lower() or 'get in touch' in s.lower()))
    if contact_links:
        contact_href = contact_links[0].get('href', '')
        if contact_href:
            contact_url = urljoin(url, contact_href)
    
    address_elements = soup.find_all(['p', 'div', 'span'], 
                                    string=lambda s: s and any(x in s.lower() for x in ['address', 'location', 'street', 'avenue', 'blvd', 'road']))
    if address_elements:
        address = address_elements[0].get_text().strip()
    
    return phone, email, contact_url, address

def extract_company_name(soup, url, page_title):
    """
    Extract the company name from the webpage content using various methods.
    
    Args:
        soup: BeautifulSoup object of the webpage
        url: URL of the webpage
        page_title: Original page title as fallback
        
    Returns:
        Extracted company name or original page title if extraction fails
    """
    # Try to get company name from og:site_name meta tag (commonly used for company name)
    og_site = soup.find('meta', {'property': 'og:site_name'})
    if og_site and og_site.get('content'):
        return og_site.get('content').strip()
    
    # Try to get from schema.org Organization or LocalBusiness structured data
    schema_org = soup.find('script', {'type': 'application/ld+json'})
    if schema_org:
        try:
            data = json.loads(schema_org.string)
            if isinstance(data, dict):
                # Check for name in schema.org data
                if data.get('@type') in ['Organization', 'LocalBusiness', 'Corporation', 'Company'] and data.get('name'):
                    return data.get('name')
                # Check for publisher name
                if data.get('publisher', {}).get('name'):
                    return data.get('publisher', {}).get('name')
        except:
            pass
    
    # Try to get from the title tag - often contains company name
    title_tag = soup.find('title')
    if title_tag and title_tag.text:
        title_text = title_tag.text.strip()
        # If title contains separator like ' - ' or ' | ', the first part is often the page name
        # and the last part is often the company name
        for separator in [' - ', ' | ', ' :: ', ' » ', ' – ']:
            if separator in title_text:
                parts = title_text.split(separator)
                if len(parts) >= 2:
                    # Usually the last part is the company name
                    return parts[-1].strip()
    
    # Try to get from common header elements
    for selector in ['header .logo', '.site-title', '#logo', '.brand', '.company-name', '.logo img', '.header .logo']:
        element = soup.select_one(selector)
        if element:
            if element.name == 'img' and element.get('alt'):
                return element.get('alt').strip()
            elif element.text:
                return element.text.strip()
    
    # If all else fails, return the original page title
    return page_title

def visit_and_check_website(url, business_name):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=15)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Extract company name from webpage content
        company_name = extract_company_name(soup, url, business_name)
        
        has_chatbot = detect_chatbot(soup)
        
        phone, email, contact_url, address = extract_contact_info(soup, url)
        
        description = ""
        meta_desc = soup.find('meta', {'name': 'description'}) or soup.find('meta', {'property': 'og:description'})
        if meta_desc:
            description = meta_desc.get('content', '').strip()
            if len(description) > 200:
                description = description[:197] + '...'
        
        location = address
        
        description = description.replace('\n', ' ').replace('\r', ' ').strip()
        location = location.replace('\n', ' ').replace('\r', ' ').strip()
        
        for _ in range(3):
            pyautogui.scroll(-500)
            time.sleep(1)
        
        return company_name, has_chatbot, phone, email, contact_url, address, description, location
    
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
    for _ in range(10):
        pyautogui.scroll(-700)
        time.sleep(0.5)
    
    next_button = locate_button("img/next_page.png", timeout=5)
    if next_button:
        pyautogui.click(next_button)
        time.sleep(5)
        return True
    
    return False

def process_search_term(search_term):
    logging.info(f"\n{'#'*50}\nStarting search for: {search_term}\n{'#'*50}")
    
    visited_urls = load_visited_urls(visited_urls_file)
    logging.info(f"Loaded {len(visited_urls)} visited URLs from {visited_urls_file}")
    main_window = open_search_engine(search_term)
    
    if not main_window:
        return
    
    if not click_places_tab():
        logging.error(f"Could not access Places tab for '{search_term}'")
        return
    
    for page in range(1, max_pages + 1):
        logging.info(f"Processing Places page {page} for '{search_term}'")
        
        processed_count = 0
        max_businesses_per_page = 10
        
        last_y_position = None
        with alive_bar(max_businesses_per_page, title=f'Places Page {page}', bar='filling', spinner='dots_waves') as bar:
            for business_idx in range(max_businesses_per_page):
                main_window.activate()
                time.sleep(1)
                
                result, last_y_position = process_business_listing(main_window, visited_urls, search_term, last_y_position)
                if result:
                    processed_count += 1
                
                bar()
                
                if business_idx % 3 == 2:
                    main_window.activate()
                    time.sleep(1)
                    scroll_down_for_more_results()
                
                time.sleep(2)
        
        logging.info(f"Processed {processed_count} businesses on page {page}")
        
        save_visited_urls(visited_urls_file, visited_urls)
        logging.info(f"Saved {len(visited_urls)} URLs to {visited_urls_file}")
        
        if page < max_pages:
            main_window.activate()
            time.sleep(1)
            if not navigate_to_next_page():
                logging.warning(f"Could not navigate to page {page+1}, stopping")
                break
    
    try:
        main_window.close()
    except:
        pass

def main():
    initialize_csv()
    
    for search_term in search_terms:
        process_search_term(search_term)
        time.sleep(5)
    
    logging.info("Script completed successfully")

if __name__ == "__main__":
    main()
