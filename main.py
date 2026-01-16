import os
import sys
import logging
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
from simple_salesforce import Salesforce
from DrissionPage import ChromiumPage, ChromiumOptions

# Secrets
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

def send_email_msg(subject, body):
    if not EMAIL_SENDER or not EMAIL_PASSWORD or not EMAIL_RECEIVER:
        logging.warning("Email secrets missing. Skipping notification.")
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

def get_sf_connection():
    try:
        sf = Salesforce(username=SF_USERNAME, password=SF_PASSWORD, security_token=SF_TOKEN)
        logging.info("Connected to Salesforce API")
        return sf
    except Exception as e:
        logging.error(f"Connection Failed: {e}")
        send_email_msg("Script Failed", f"Salesforce Connection Error: {e}")
        sys.exit(1)

# ... [Helpers same as before] ...
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

def scrape_record(browser, rec_id, obj_type):
    url = BASE_URL.format(obj=obj_type, id=rec_id)
    try:
        browser.get(url)
        browser.wait.doc_loaded(timeout=15)
        try:
            if browser.ele('xpath://button[contains(text(), "Show All")]', timeout=3):
                browser.ele('xpath://button[contains(text(), "Show All")]').click()
            if browser.ele('css:button.testonly-expandAll', timeout=3):
                browser.ele('css:button.testonly-expandAll').click()
            for btn in browser.eles('xpath://button[contains(., "Repl")]', timeout=3):
                if btn.is_displayed and "Collapse" not in btn.text:
                    browser.driver.execute_script("arguments[0].click();", btn)
        except: pass
        browser.scroll.to_bottom(); browser.wait(1)
        
        cutoff_y = 0
        markers = browser.eles('css:.slds-timeline__date', timeout=2)
        if markers: cutoff_y = markers[0].rect.location['y']
        elif browser.ele('xpath://span[contains(text(), "Upcoming & Overdue")]', timeout=1): cutoff_y = 999999

        valid_dates = []
        for d in browser.eles('css:.dueDate', timeout=5):
            if 'overdue' in d.text.lower(): continue
            if cutoff_y > 0 and d.rect.location['y'] < cutoff_y: continue
            cl = clean_activity_date(d.text)
            if cl: valid_dates.append(cl)
        return len(valid_dates), (valid_dates[0] if valid_dates else None)
    except: return 0, None

def main():
    sf = get_sf_connection()
    
    # === UPDATED BROWSER SETTINGS (CRITICAL FOR GITHUB) ===
    co = ChromiumOptions()
    co.set_argument('--headless=new')
    co.set_argument('--no-sandbox')
    # Ye 2 lines nayi hain (Crash rokne ke liye):
    co.set_argument('--disable-gpu')
    co.set_argument('--disable-dev-shm-usage') 
    
    try:
        browser = ChromiumPage(addr_or_opts=co)
    except Exception as e:
        logging.error(f"Browser Launch Failed: {e}")
        send_email_msg("Browser Error", str(e))
        sys.exit(1)
    
    # Auto Login
    domain = BASE_URL.split('/')[2]
    browser.get(f"https://{domain}/secur/frontdoor.jsp?sid={sf.session_id}")
    browser.wait.doc_loaded()

    # Queries
    start_dt, end_dt = get_this_week_soql_filter()
    mkt_query = f"SELECT Id FROM Lead WHERE LeadSource = 'Marketing Inbound' AND CreatedDate >= {start_dt} AND CreatedDate <= {end_dt}"
    target_owners = "('Harshit Gupta', 'Abhishek Nayak', 'Deepesh Dubey', 'Prashant Jha')"
    sales_query = f"SELECT Id, Owner.Name FROM Account WHERE Owner.Name IN {target_owners}"

    try:
        mkt_recs = sf.query_all(mkt_query)['records']
        sales_recs = sf.query_all(sales_query)['records']
        start_msg = f"🚀 Started updating {len(mkt_recs)} Marketing Leads and {len(sales_recs)} Sales Accounts."
        send_email_msg("Salesforce Bot Started", start_msg)
    except Exception as e:
        send_email_msg("Error Fetching Data", str(e))
        browser.quit(); sys.exit(1)

    mkt_success = 0
    logging.info(f"Processing {len(mkt_recs)} Leads...")
    for rec in mkt_recs:
        lid = rec['Id']
        count, last_date = scrape_record(browser, lid, 'Lead')
        try:
            payload = {MKT_API_COUNT: count}
            if api_date := convert_date_for_api(last_date): payload[MKT_API_DATE] = api_date
            sf.Lead.update(lid, payload)
            mkt_success += 1
        except: pass

    sales_success = 0
    logging.info(f"Processing {len(sales_recs)} Accounts...")
    for rec in sales_recs:
        aid = rec['Id']
        count, last_date = scrape_record(browser, aid, 'Account')
        try:
            if api_date := convert_date_for_api(last_date):
                sf.Account.update(aid, {SALES_API_DATE: api_date})
                sales_success += 1
        except: pass

    browser.quit()
    end_msg = f"✅ Update complete: {mkt_success} Leads and {sales_success} Accounts updated successfully."
    send_email_msg("Salesforce Bot Success", end_msg)

if __name__ == "__main__":
    main()
