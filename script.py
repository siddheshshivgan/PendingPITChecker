from datetime import datetime,timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import glob
import os
from pathlib import Path
import smtplib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import pytesseract
import time
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv(override=True)
# initialize start and end dates
end_date = datetime.now()
start_date = end_date - timedelta(days=7)

# Get the home directory
home_dir = Path.home()

# Construct the path to the Downloads folder
downloads_dir = home_dir / "Downloads"

# Set up your Tesseract OCR path if it's not in your PATH environment variable
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

chrome_options = Options()
# chrome_options.add_argument("--headless")
# chrome_options.add_argument("--no-sandbox")
# chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")  # Ensure a standard window size

# Create a new instance of the Chrome driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Function to format pending transactions for email
def format_pending_email(all_pending_transactions):
    """
    Format pending transactions into HTML email body
    """
    if not all_pending_transactions:
        return None
    
    html = """
    <html>
    <head>
        <style>
            table {
                border-collapse: collapse;
                width: 100%;
                font-family: Arial, sans-serif;
            }
            th {
                background-color: #4CAF50;
                color: white;
                padding: 12px;
                text-align: left;
                border: 1px solid #ddd;
            }
            td {
                padding: 10px;
                border: 1px solid #ddd;
            }
            tr:nth-child(even) {
                background-color: #f2f2f2;
            }
            h2 {
                color: #d9534f;
            }
            .summary {
                background-color: #fff3cd;
                padding: 15px;
                margin-bottom: 20px;
                border-left: 4px solid #ffc107;
            }
        </style>
    </head>
    <body>
        <h2>⚠️ Pending Transactions Alert - """ + datetime.now().strftime('%d-%m-%Y') + """</h2>
        <div class="summary">
            <strong>Total Pending Transactions: """ + str(len(all_pending_transactions)) + """</strong>
        </div>
        <table>
            <tr>
                <th>Account</th>
                <th>TAR No</th>
                <th>Investor</th>
                <th>Scheme</th>
                <th>Type</th>
                <th>Amount</th>
                <th>Bank</th>
                <th>Status</th>
            </tr>
    """
    
    for txn in all_pending_transactions:
        html += f"""
            <tr>
                <td>{txn['Account']}</td>
                <td>{txn['TAR_No']}</td>
                <td>{txn['Investor']}</td>
                <td>{txn['Scheme']}</td>
                <td>{txn['Transaction_Type']}</td>
                <td>{txn['Investment_Amount']}</td>
                <td>{txn['Bank']}</td>
                <td style="color: red; font-weight: bold;">{txn['Auth_Status']}</td>
            </tr>
        """
    
    html += """
        </table>
        <br>
        <p style="color: #666;">
            <em>Please authorize these transactions at your earliest convenience.</em>
        </p>
    </body>
    </html>
    """
    
    return html

# Function to send email
def send_email(to_address, subject, body):
    from_address = os.environ.get('EMAIL_ID')
    app_password = os.environ.get('PASSWORD')

    # Set up the server
    server = smtplib.SMTP(host='smtp.gmail.com', port=587)
    server.starttls()
    server.login(from_address, app_password)

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = formataddr(('Shivgan Associates', from_address))
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    # Send the email
    server.send_message(msg)
    server.quit()

# Function to check pending transactions
def check_pending_transactions(page_source, account_name):
    """
    Check for today's pending transactions in the page source
    Returns list of pending transactions
    """
    soup = BeautifulSoup(page_source, 'html.parser')
    today = datetime.now().strftime('%d-%m-%Y')
    
    pending_transactions = []
    
    # Find all transaction rows
    rows = soup.find_all('tr', class_=['ev_dhx_web', 'odd_dhx_web'])
    
    for row in rows:
        cells = row.find_all('td')
        if len(cells) > 23:  # Ensure row has enough columns
            try:
                submission_date = cells[3].get_text(strip=True)
                auth_status = cells[23].get_text(strip=True)
                
                # Check if it's today's date and status is PENDING
                if submission_date == today and auth_status == 'PENDING':
                    transaction = {
                        'Account': account_name,
                        'Sr_No': cells[0].get_text(strip=True),
                        'Submission_Date': submission_date,
                        'TAR_No': cells[4].get_text(strip=True),
                        'Partner': cells[6].get_text(strip=True),
                        'Group': cells[7].get_text(strip=True),
                        'Investor': cells[8].get_text(strip=True),
                        'Client_Code': cells[9].get_text(strip=True),
                        'Scheme': cells[11].get_text(strip=True),
                        'Transaction_Type': cells[14].get_text(strip=True),
                        'Investment_Amount': cells[15].get_text(strip=True),
                        'Bank': cells[18].get_text(strip=True),
                        'Auth_Mode': cells[19].get_text(strip=True),
                        'Auth_Status': auth_status,
                    }
                    pending_transactions.append(transaction)
            except Exception as e:
                continue
    
    return pending_transactions
    
# Function to get xls file paths
def get_latest_xls_files(num_files=3):
    # Construct the path to the Downloads folder
    downloads_path = os.path.expanduser(downloads_dir)

    # Search for .xls files in the Downloads folder
    search_pattern = os.path.join(downloads_path, '*.xls')
    xls_files = glob.glob(search_pattern)

    # Sort files by modification time (latest first)
    xls_files.sort(key=os.path.getmtime, reverse=True)

    # Take the first num_files files
    latest_xls_files = xls_files[:num_files]

    return latest_xls_files

# Function to login
def login(user_id,pwd):
    # Locate the username and password fields and enter the login details
    username = driver.find_element(By.NAME, 'partnerId1')
    password = driver.find_element(By.NAME, 'password1')
    username.send_keys(user_id)
    password.send_keys(pwd)
    print(user_id)
    
    # Capture the CAPTCHA image
    captcha_image = driver.find_element(By.ID, 'imgCaptcha')  # Update the XPath

    # Save the CAPTCHA image
    captcha_image.screenshot('captcha.png')

    # Use OCR to read the CAPTCHA
    captcha_text = pytesseract.image_to_string(Image.open('captcha.png')).strip()
    captcha_text = captcha_text.replace(" ", "")

    # Enter the CAPTCHA text
    captcha_field = driver.find_element(By.NAME, 'capcode')
    captcha_field.send_keys(captcha_text)

    # Submit the form
    driver.find_element(By.NAME, 'action').click()
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'action'))).click()

accounts = [
    {
        "name": "SID",
        "id": os.environ.get('SID_ID'),
        "password": os.environ.get('SID_PASSWORD')
    },
    {
        "name": "RAJAN",
        "id": os.environ.get('RAJAN_ID'),
        "password": os.environ.get('RAJAN_PASSWORD')
    },
    {
        "name": "RESHMA",
        "id": os.environ.get('RESHMA_ID'),
        "password": os.environ.get('RESHMA_PASSWORD')
    }
]

# Navigate to the login page
driver.get(os.environ.get('PARTNER_DESK'))

# Store all pending transactions from all accounts
all_pending_transactions = []

for acc in accounts:
    login(acc['id'],acc['password'])
    time.sleep(10)

    # check if captcha failed
    if 'E-MF Account' not in driver.page_source:
        login(acc['id'],acc['password'])
        time.sleep(10)

    if 'popupCloseButton' in driver.page_source:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'popupCloseButton'))).click()

    
    time.sleep(2)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@onclick="javascript:closeDiwaliSIPPopup();"]'))).click()
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//a[text()="Stock Exchange"]'))).click()
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//b[text()="Transaction Authorization Report"]'))).click()
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 'apply'))).click()
    time.sleep(5)
    
    # Get the page source after the report loads
    page_source = driver.page_source
    
    # Check for pending transactions
    pending_txns = check_pending_transactions(page_source, acc['name'])
    
    if pending_txns:
        print(f"\n{'='*80}")
        print(f"PENDING TRANSACTIONS FOUND FOR {acc['name']}")
        print(f"{'='*80}")
        for txn in pending_txns:
            print(f"\nTAR No: {txn['TAR_No']}")
            print(f"Investor: {txn['Investor']}")
            print(f"Scheme: {txn['Scheme']}")
            print(f"Type: {txn['Transaction_Type']}")
            print(f"Amount: {txn['Investment_Amount']}")
            print(f"Status: {txn['Auth_Status']}")
        
        # Add to all pending transactions
        all_pending_transactions.extend(pending_txns)
    else:
        print(f"\n✓ No pending transactions for {acc['name']}")
    

    
       
    driver.get(os.environ.get('PARTNER_DESK'))

# Send email if there are pending transactions
if all_pending_transactions:
    email_body = format_pending_email(all_pending_transactions)
    
    # Email recipients (modify as needed)
    recipients = os.environ.get('SID_EMAIL_ID') + ',' + os.environ.get('RAJAN_EMAIL_ID')
    
    try:
        send_email(
            to_address=recipients,
            subject=f"⚠️ Pending Transactions Alert - {datetime.now().strftime('%d-%m-%Y')}",
            body=email_body
        )
        print(f"\n✓ Email alert sent successfully to {recipients}")
    except Exception as e:
        print(f"\n✗ Failed to send email: {str(e)}")
    
    print(f"\n{'='*80}")
    print(f"TOTAL PENDING TRANSACTIONS: {len(all_pending_transactions)}")
    print(f"{'='*80}")
else:
    print(f"\n{'='*80}")
    print("✓ NO PENDING TRANSACTIONS FOUND FOR TODAY")
    print(f"{'='*80}")

driver.quit()

