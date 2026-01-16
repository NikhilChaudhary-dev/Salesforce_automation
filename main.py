import os
import sys
import time
import logging
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid, formatdate
from datetime import datetime, timedelta
from collections import Counter
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
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER') # Multiple emails comma separated

BASE_URL = 'https://loop-subscriptions.lightning.force.com/lightning/r/{obj}/{id}/view'
MKT_API_COUNT = 'Count_of_Activities__c'
MKT_API_DATE  = 'Last_Activity_Date__c'
SALES_API_DATE = 'Last_Activity_Date_V__c'

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s')

# ================= HELPER: GET INDIA DATE =================
def get_india_date_str():
    # Server UTC par hota hai, +5:30 add karke India Time layenge
    utc_now = datetime.utcnow()
    ist_now = utc_now + timedelta(hours=5, minutes=30)
    return ist_now.strftime('%d-%b-%Y (India Time)')

# ================= EMAIL THREADING FUNCTION =================
def send_email_thread(subject, body, parent_msg_id=None):
    if not EMAIL_SENDER or not EMAIL_PASSWORD or not EMAIL_RECEIVER:
        logging.warning("Email secrets missing. Skipping.")
        return None

    try:
        msg = EmailMessage()
        msg.set_content(body)
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER # Handles "a@b.com, c@d.com"
        msg['Date'] = formatdate(localtime=True)
        
        # Unique ID generate karo is email ke liye
        new_msg_id = make_msgid()
        msg['Message-ID'] = new_msg_id

        # THREADING LOGIC
        if parent_msg_id:
            # Agar ye reply hai, to Subject me Re: lagao aur Headers set karo
            msg['Subject'] = f"Re: {subject}"
            msg['In-Reply-To'] = parent_msg_id
            msg['References'] = parent_msg_id
        else:
            # Ye naya email hai
            msg['Subject'] = subject

        # Send Mail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        
        logging.info(f"Email sent. ID: {new_msg_id}")
        return new_msg_id # Is ID ko return karo taaki agla mail ispe reply kare

    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return None

# ================= SALESFORCE & BROWSER =================
def get_sf_connection():
    try:
        sf = Salesforce(username=SF_USERNAME, password=SF_PASSWORD, security_token=SF_TOKEN)
        logging.info("Connected to Salesforce API")
        return sf
    except Exception as e:
        sys.exit(1)

def get_selenium_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

# ================= DATE UTILS =================
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

# ================= SCRAPING LOGIC =================
def scrape_record(driver, rec_id, obj_type):
    url = BASE_URL.format(obj=obj_type, id=rec_id)
    try:
        driver.get(url)
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".slds-timeline__item")))
        except: pass

        try:
            buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Show All') or contains(@class, 'testonly-expandAll')]")
            for btn in buttons: driver.execute_script("arguments[0].click();", btn); time.sleep(1)
            
            reply_buttons = driver.find_elements(By.XPATH, "//button[contains(., 'Repl')]")
            for btn in reply_buttons: 
                if btn.is_displayed(): driver.execute_script("arguments[0].click();", btn)
        except: pass
        time.sleep(2)

        cutoff_y = 0
        try:
            markers = driver.find_elements(By.CSS_SELECTOR, ".slds-timeline__date")
            if markers: cutoff_y = markers[0].location['y']
            else:
                upcoming_text = driver.find_elements(By.XPATH, "//span[contains(text(), 'Upcoming & Overdue')]")
                if upcoming_text: cutoff_y = 999999
        except: pass

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        valid_dates = []
        for el in soup.select('.dueDate'):
            # Convert BS4 element back to check logic roughly or iterate manually
            pass 
        
        # Using Selenium logic for position filtering
        date_elements = driver.find_elements(By.CSS_SELECTOR, ".dueDate")
        for el in date_elements:
            text = el.text.strip()
            if 'overdue' in text.lower(): continue
            if cutoff_y > 0 and el.location['y'] < cutoff_y: continue
            cl = clean_activity_date(text)
            if cl: valid_dates.append(cl)

        return len(valid_dates), (valid_dates[0] if valid_dates else None)
    except: return 0, None

# ================= MAIN =================
def main():
    sf = get_sf_connection()
    
    # 1. Queries
    start_dt, end_dt = get_this_week_soql_filter()
    mkt_query = f"SELECT Id FROM Lead WHERE LeadSource = 'Marketing Inbound' AND CreatedDate >= {start_dt} AND CreatedDate <= {end_dt}"
    target_owners = "('Harshit Gupta', 'Abhishek Nayak', 'Deepesh Dubey', 'Prashant Jha')"
    sales_query = f"SELECT Id, Owner.Name FROM Account WHERE Owner.Name IN {target_owners}"

    try:
        mkt_recs = sf.query_all(mkt_query)['records']
        sales_recs = sf.query_all(sales_query)['records']
        
        # --- BREAKDOWN LOGIC ---
        # Sales Breakdown calculate karna
        sales_counts = Counter([r['Owner']['Name'] for r in sales_recs])
        sales_breakdown_str = "\n".join([f"   - {owner}: {count} Accounts" for owner, count in sales_counts.items()])
        
        india_date = get_india_date_str()
        
        # --- START EMAIL BODY ---
        start_subject = f"Salesforce Updates - {india_date}"
        start_body = (
            f"🚀 Started updating {len(mkt_recs)} Marketing Leads and {len(sales_recs)} Sales Accounts.\n"
            f"📅 Date: {india_date}\n\n"
            f"📊 **Details Summary:**\n"
            f"👉 **Marketing Inbound:** Total {len(mkt_recs)} Leads mili hain.\n\n"
            f"👉 **Sales Breakdown (Accounts found):**\n"
            f"{sales_breakdown_str}\n\n"
            f"Updating now...\n\n"
            f"Regards,\nHappy Bot 🤖"
        )
        
        # SEND START EMAIL & SAVE THREAD ID
        thread_id = send_email_thread(start_subject, start_body, parent_msg_id=None)

    except Exception as e:
        # Agar fail hua to new mail bhej do
        send_email_thread("Script Failed", str(e))
        sys.exit(1)

    # 2. Browser Start
    try:
        driver = get_selenium_driver()
        domain = BASE_URL.split('/')[2]
        driver.get(f"https://{domain}/secur/frontdoor.jsp?sid={sf.session_id}")
        time.sleep(5)
    except Exception as e:
        driver.quit(); sys.exit(1)

    # 3. Processing
    mkt_success = 0
    for rec in mkt_recs:
        lid = rec['Id']
        count, last_date = scrape_record(driver, lid, 'Lead')
        try:
            payload = {MKT_API_COUNT: count}
            if api_date := convert_date_for_api(last_date): payload[MKT_API_DATE] = api_date
            sf.Lead.update(lid, payload)
            mkt_success += 1
        except: pass

    sales_success = 0
    for rec in sales_recs:
        aid = rec['Id']
        count, last_date = scrape_record(driver, aid, 'Account')
        try:
            if api_date := convert_date_for_api(last_date):
                sf.Account.update(aid, {SALES_API_DATE: api_date})
                sales_success += 1
        except: pass

    driver.quit()

    # --- END EMAIL (THREADED) ---
    end_body = (
        f"✅ Update complete: Same data updated successfully.\n\n"
        f"📊 **Final Status:**\n"
        f"   - Marketing Leads Updated: {mkt_success}/{len(mkt_recs)}\n"
        f"   - Sales Accounts Updated: {sales_success}/{len(sales_recs)}\n\n"
        f"Regards,\nHappy Bot 🤖"
    )
    
    # Send Reply to the same thread
    send_email_thread(start_subject, end_body, parent_msg_id=thread_id)

if __name__ == "__main__":
    main()
