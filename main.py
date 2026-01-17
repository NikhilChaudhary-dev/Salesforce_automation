import os
import sys
import time
import logging
import smtplib
import csv
import io
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
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER')

BASE_URL = 'https://loop-subscriptions.lightning.force.com/lightning/r/{obj}/{id}/view'
MKT_API_COUNT = 'Count_of_Activities__c'
MKT_API_DATE  = 'Last_Activity_Date__c'
SALES_API_DATE = 'Last_Activity_Date_V__c'

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s')

# ================= HELPER: GET INDIA DATE =================
def get_india_date_str():
    utc_now = datetime.utcnow()
    ist_now = utc_now + timedelta(hours=5, minutes=30)
    return ist_now.strftime('%d-%b-%Y (IST)')

# ================= HTML EMAIL TEMPLATE (PERSONALIZED) =================
def create_html_body(title, data_rows, footer_note=""):
    rows_html = ""
    for label, value in data_rows:
        rows_html += f"""
        <tr>
            <td style="padding: 12px; border-bottom: 1px solid #e0e0e0; font-weight: bold; color: #333; width: 40%;">{label}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e0e0e0; color: #555;">{value}</td>
        </tr>
        """
    
    html = f"""
    <html>
    <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; line-height: 1.6; background-color: #f9f9f9; padding: 20px;">
        <div style="max-width: 600px; margin: auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
            
            <h2 style="color: #2c3e50; margin-top: 0; border-bottom: 2px solid #3498db; padding-bottom: 10px;">
                {title}
            </h2>
            <p style="font-size: 14px; color: #7f8c8d; margin-bottom: 20px;">{get_india_date_str()}</p>
            
            <table style="width: 100%; border-collapse: collapse;">
                {rows_html}
            </table>
            
            <p style="margin-top: 25px; font-style: italic; color: #7f8c8d; font-size: 13px;">{footer_note}</p>
            
            <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 15px; font-size: 12px; color: #999; text-align: center;">
                Automated by <b>Nikhil Chaudhary</b> ⚡
            </div>
        </div>
    </body>
    </html>
    """
    return html

# ================= EMAIL THREADING FUNCTION =================
def send_email_thread(subject, html_content, parent_msg_id=None, csv_data=None):
    if not EMAIL_SENDER or not EMAIL_PASSWORD or not EMAIL_RECEIVER:
        logging.warning("Email secrets missing. Skipping notification.")
        return None

    try:
        msg = EmailMessage()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Date'] = formatdate(localtime=True)
        
        new_msg_id = make_msgid()
        msg['Message-ID'] = new_msg_id

        if parent_msg_id:
            msg['Subject'] = f"Re: {subject}"
            msg['In-Reply-To'] = parent_msg_id
            msg['References'] = parent_msg_id
        else:
            msg['Subject'] = subject

        msg.set_content("Please enable HTML to view this report.")
        msg.add_alternative(html_content, subtype='html')
        
        if csv_data:
            msg.add_attachment(csv_data.encode('utf-8'), maintype='text', subtype='csv', filename='error_report.csv')

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        
        logging.info(f"Email sent successfully. ID: {new_msg_id}")
        return new_msg_id

    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return None

# ================= SALESFORCE & BROWSER =================
def get_sf_connection():
    try:
        sf = Salesforce(username=SF_USERNAME, password=SF_PASSWORD, security_token=SF_TOKEN)
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

# ================= UTILS =================
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
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".slds-timeline__item")))
        except: pass

        try:
            buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Show All') or contains(@class, 'testonly-expandAll')]")
            for btn in buttons: driver.execute_script("arguments[0].click();", btn); time.sleep(0.5)
            reply_buttons = driver.find_elements(By.XPATH, "//button[contains(., 'Repl')]")
            for btn in reply_buttons: 
                if btn.is_displayed(): driver.execute_script("arguments[0].click();", btn)
        except: pass
        time.sleep(1.5)

        cutoff_y = 0
        try:
            markers = driver.find_elements(By.CSS_SELECTOR, ".slds-timeline__date")
            if markers: cutoff_y = markers[0].location['y']
            else:
                upcoming_text = driver.find_elements(By.XPATH, "//span[contains(text(), 'Upcoming & Overdue')]")
                if upcoming_text: cutoff_y = 999999
        except: pass

        valid_dates = []
        date_elements = driver.find_elements(By.CSS_SELECTOR, ".dueDate")
        for el in date_elements:
            text = el.text.strip()
            if 'overdue' in text.lower(): continue
            if cutoff_y > 0 and el.location['y'] < cutoff_y: continue
            cl = clean_activity_date(text)
            if cl: valid_dates.append(cl)

        return len(valid_dates), (valid_dates[0] if valid_dates else None)
    except Exception as e:
        raise e

# ================= MAIN EXECUTION =================
def main():
    sf = get_sf_connection()
    failed_records_log = []

    # 1. Prepare Queries
    start_dt, end_dt = get_this_week_soql_filter()
    mkt_query = f"SELECT Id FROM Lead WHERE LeadSource = 'Marketing Inbound' AND CreatedDate >= {start_dt} AND CreatedDate <= {end_dt}"
    target_owners = "('Harshit Gupta', 'Abhishek Nayak', 'Deepesh Dubey', 'Prashant Jha')"
    sales_query = f"SELECT Id, Owner.Name FROM Account WHERE Owner.Name IN {target_owners}"

    try:
        mkt_recs = sf.query_all(mkt_query)['records']
        sales_recs = sf.query_all(sales_query)['records']
        
        sales_counts = Counter([r['Owner']['Name'] for r in sales_recs])
        # Nicely formatted bullets for HTML
        sales_breakdown = "<br>".join([f"• {owner}: <b>{count}</b>" for owner, count in sales_counts.items()])
        
        # --- TITLE CHANGE HERE ---
        title = "📊 Salesforce Daily Activity Report"
        
        data = [
            ("Date", get_india_date_str()),
            ("Marketing Inbound Leads Found", f"{len(mkt_recs)} Leads"),
            ("Sales Accounts Found", f"{len(sales_recs)} Accounts"),
            ("Sales Breakdown", sales_breakdown)
        ]
        html_body = create_html_body(title, data, "The automation script has started. You will receive a summary upon completion.")
        thread_id = send_email_thread(title, html_body)

    except Exception as e:
        send_email_thread("Script Failed", f"<p>Critical Error: {str(e)}</p>")
        sys.exit(1)

    # 2. Browser Start
    try:
        driver = get_selenium_driver()
        domain = BASE_URL.split('/')[2]
        driver.get(f"https://{domain}/secur/frontdoor.jsp?sid={sf.session_id}")
        time.sleep(5)
    except Exception as e:
        driver.quit(); sys.exit(1)

    # 3. Process Marketing Leads
    mkt_stats = {'updated': 0, 'skipped': 0, 'failed': 0}
    for rec in mkt_recs:
        lid = rec['Id']
        try:
            count, last_date = scrape_record(driver, lid, 'Lead')
            if last_date:
                payload = {MKT_API_COUNT: count}
                if api_date := convert_date_for_api(last_date): payload[MKT_API_DATE] = api_date
                sf.Lead.update(lid, payload)
                mkt_stats['updated'] += 1
            else:
                mkt_stats['skipped'] += 1
        except Exception as e: 
            mkt_stats['failed'] += 1
            failed_records_log.append(['Lead', lid, str(e)])

    # 4. Process Sales Accounts
    sales_stats = {'updated': 0, 'skipped': 0, 'failed': 0}
    for rec in sales_recs:
        aid = rec['Id']
        try:
            count, last_date = scrape_record(driver, aid, 'Account')
            if last_date:
                if api_date := convert_date_for_api(last_date):
                    sf.Account.update(aid, {SALES_API_DATE: api_date})
                    sales_stats['updated'] += 1
                else:
                    sales_stats['skipped'] += 1
            else:
                sales_stats['skipped'] += 1
        except Exception as e: 
            sales_stats['failed'] += 1
            failed_records_log.append(['Account', aid, str(e)])

    driver.quit()

    # --- CSV LOGIC ---
    csv_string = None
    footer_note = "Note: 'Skipped' records were checked but had no valid activity date."
    if failed_records_log:
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(['Type', 'Record ID', 'Error Reason'])
        writer.writerows(failed_records_log)
        csv_string = output.getvalue()
        footer_note += " ⚠️ <b>Errors detected. Please check the attached CSV.</b>"

    # --- COMPLETION EMAIL ---
    end_title = "✅ Execution Complete"
    
    mkt_result = (f"<b>{mkt_stats['updated']}</b> Updated<br>"
                  f"<span style='color:#f39c12;'>{mkt_stats['skipped']} Skipped</span><br>"
                  f"<span style='color:#c0392b;'>{mkt_stats['failed']} Failed</span>")
    
    sales_result = (f"<b>{sales_stats['updated']}</b> Updated<br>"
                    f"<span style='color:#f39c12;'>{sales_stats['skipped']} Skipped</span><br>"
                    f"<span style='color:#c0392b;'>{sales_stats['failed']} Failed</span>")

    end_data = [
        ("Final Status", "Success" if not failed_records_log else "Completed with Errors"),
        ("Marketing Leads", mkt_result),
        ("Sales Accounts", sales_result),
        ("Total Records Processed", f"{len(mkt_recs) + len(sales_recs)}")
    ]
    
    html_body = create_html_body(end_title, end_data, footer_note)
    
    send_email_thread(title, html_body, parent_msg_id=thread_id, csv_data=csv_string)

if __name__ == "__main__":
    main()
