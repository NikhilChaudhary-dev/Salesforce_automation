import os
import sys
import time
import logging
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
from simple_salesforce import Salesforce
from bs4 import BeautifulSoup

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIGURATION =================
SF_USERNAME = os.getenv('SF_USERNAME')
SF_PASSWORD = os.getenv('SF_PASSWORD')
SF_TOKEN    = os.getenv('SF_TOKEN')

# Email Config
EMAIL_SENDER   = os.getenv('EMAIL_SENDER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER')

BASE_URL = 'https://loop-subscriptions.lightning.force.com/lightning/r/{obj}/{id}/view'
MKT_API_COUNT = 'Count_of_Activities__c'
MKT_API_DATE  = 'Last_Activity_Date__c'
SALES_API_DATE = 'Last_Activity_Date_V__c'

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s')

# ================= EMAIL HELPER =================
def send_email_msg(subject, body):
    if not EMAIL_SENDER or not EMAIL_PASSWORD or not EMAIL_RECEIVER:
        logging.warning("Email secrets missing. Skipping.")
        return
    try:
        msg = EmailMessage()
        msg.set_content(body)
        msg['Subject'] = subject
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        logging.info(f"Email sent to {EMAIL_RECEIVER}")
    except Exception as e:
        logging.error(f"Failed to send email: {e}")

# ================= SALESFORCE CONNECTION =================
def get_sf_connection():
    try:
        sf = Salesforce(username=SF_USERNAME, password=SF_PASSWORD, security_token=SF_TOKEN)
        logging.info("Connected to Salesforce API")
        return sf
    except Exception as e:
        logging.error(f"Connection Failed: {e}")
        send_email_msg("Script Failed", f"Salesforce Connection Error: {e}")
        sys.exit(1)

# ================= DATE HELPERS =================
def get_this_week_soql_filter():
    today = datetime.now()
    start_of_week = today - timedelta(days=today.weekday())
    start_of_week = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_week = start_of_week + timedelta(days=6, hours=23, minutes=59, seconds=59)
    return start_of_week.strftime('%Y-%m-%dT00:00:00Z'), end_of_week.strftime('%Y-%m-%dT23:59:59Z')

def clean_activity_date(text):
    if not text: return ""
    text = text.split('|')[-1].strip()
    text_lower = text.lower()
    now = datetime.now()
    if 'today' in text_lower: return now.strftime('%d-%b-%Y')
    elif 'yesterday' in text_lower: return (now - timedelta(days=1)).strftime('%d-%b-%Y')
    elif 'tomorrow' in text_lower: return (now + timedelta(days=1)).strftime('%d-%b-%Y')
    if 'overdue' in text_lower: text = text_lower.replace('overdue', '').strip().title()
    return text

def convert_date_for_api(date_str):
    if not date_str: return None
    try: return datetime.strptime(date_str, '%d-%b-%Y').strftime('%Y-%m-%d')
    except: return None

# ================= BROWSER SETUP (SELENIUM) =================
def get_selenium_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new") # Important for Server
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # Auto-install ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

# ================= SCRAPING LOGIC =================
def scrape_record(driver, rec_id, obj_type):
    url = BASE_URL.format(obj=obj_type, id=rec_id)
    try:
        driver.get(url)
        
        # Wait for page load (Timeline element to appear)
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".slds-timeline__item"))
            )
        except:
            # Sometimes timeline is empty, that's fine
            pass

        # 1. Expand "Show All" or "View More" buttons using JS
        try:
            buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Show All') or contains(@class, 'testonly-expandAll')]")
            for btn in buttons:
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
        except: pass

        # 2. Expand "Replies"
        try:
            reply_buttons = driver.find_elements(By.XPATH, "//button[contains(., 'Repl')]")
            for btn in reply_buttons:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
        except: pass

        time.sleep(2) # Allow JS to render expansion

        # 3. Parse with BeautifulSoup (Faster & Easier)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        # Determine cutoff Y position (roughly) to filter future dates
        # Note: In BS4 we don't have Y-coordinates easily. 
        # Logic: If we see "Upcoming & Overdue", we skip dates inside that section?
        # Better approach for Selenium+BS4: Collect all valid dates and filter by Python logic if needed.
        
        valid_dates = []
        
        # Find all Due Dates
        date_elements = soup.select('.dueDate')
        
        for el in date_elements:
            text = el.get_text(strip=True)
            if 'overdue' in text.lower(): continue
            
            # Simple check: If inside "Upcoming", usually SF structures it differently. 
            # For now, let's grab all actual activity dates.
            cl = clean_activity_date(text)
            if cl: valid_dates.append(cl)

        return len(valid_dates), (valid_dates[0] if valid_dates else None)

    except Exception as e:
        logging.error(f"Error scraping {rec_id}: {e}")
        return 0, None

# ================= MAIN =================
def main():
    # 1. Connect API
    sf = get_sf_connection()
    
    # 2. Start Selenium
    try:
        driver = get_selenium_driver()
        
        # 3. Login using Session ID (Magic Trick)
        domain = BASE_URL.split('/')[2]
        frontdoor_url = f"https://{domain}/secur/frontdoor.jsp?sid={sf.session_id}"
        driver.get(frontdoor_url)
        time.sleep(5) # Wait for redirect
        logging.info("Browser Logged in via Session ID")
        
    except Exception as e:
        logging.error(f"Browser Init Failed: {e}")
        send_email_msg("Browser Error", str(e))
        sys.exit(1)

    # 4. Fetch Data
    start_dt, end_dt = get_this_week_soql_filter()
    mkt_query = f"SELECT Id FROM Lead WHERE LeadSource = 'Marketing Inbound' AND CreatedDate >= {start_dt} AND CreatedDate <= {end_dt}"
    target_owners = "('Harshit Gupta', 'Abhishek Nayak', 'Deepesh Dubey', 'Prashant Jha')"
    sales_query = f"SELECT Id, Owner.Name FROM Account WHERE Owner.Name IN {target_owners}"

    try:
        mkt_recs = sf.query_all(mkt_query)['records']
        sales_recs = sf.query_all(sales_query)['records']
        send_email_msg("Salesforce Bot Started", f"Updating {len(mkt_recs)} Leads and {len(sales_recs)} Accounts.")
    except Exception as e:
        send_email_msg("Error Fetching Data", str(e))
        driver.quit(); sys.exit(1)

    # 5. Process Marketing
    mkt_success = 0
    logging.info(f"Processing {len(mkt_recs)} Leads...")
    for rec in mkt_recs:
        lid = rec['Id']
        count, last_date = scrape_record(driver, lid, 'Lead')
        try:
            payload = {MKT_API_COUNT: count}
            if api_date := convert_date_for_api(last_date): payload[MKT_API_DATE] = api_date
            sf.Lead.update(lid, payload)
            mkt_success += 1
        except: pass

    # 6. Process Sales
    sales_success = 0
    logging.info(f"Processing {len(sales_recs)} Accounts...")
    for rec in sales_recs:
        aid = rec['Id']
        count, last_date = scrape_record(driver, aid, 'Account')
        try:
            if api_date := convert_date_for_api(last_date):
                sf.Account.update(aid, {SALES_API_DATE: api_date})
                sales_success += 1
        except: pass

    driver.quit()
    send_email_msg("Salesforce Bot Success", f"Updated {mkt_success} Leads and {sales_success} Accounts.")

if __name__ == "__main__":
    main()
