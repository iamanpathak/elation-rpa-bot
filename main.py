import time
import logging
import platform
import os
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pyautogui
import subprocess
from skimage import feature, transform, util
from PIL import Image
from scipy import ndimage
import numpy as np
import pandas as pd
import json
import csv
from selenium.webdriver.common.action_chains import ActionChains
from utils.web_driver import WebDriverManager
from config.settings import Settings
from document_type_detector import main as get_document_type
class ElationBot:
"""Cross-platform RPA Bot for Elation EMR automation"""
def __init__(self, headless=False):
self.settings = Settings()
self.web_driver_manager = WebDriverManager(
headless=headless or self.settings.HEADLESS_MODE,
window_size=self.settings.WINDOW_SIZE,
download_path=self.settings.DOWNLOAD_PATH
)
self.driver = None
self.wait = None
# Store the attributed to name from patient search
self.attributed_to_name = None
# Detect platform
self.platform = platform.system().lower()
self.is_windows = self.platform == 'windows'
self.is_macos = self.platform == 'darwin'
self.is_linux = self.platform == 'linux'
# Setup logging
logging.basicConfig(
level=logging.INFO,
format='%(asctime)s - %(levelname)s - %(message)s'
)
self.logger = logging.getLogger(__name__)
# Configure PyAutoGUI
pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = True
self.logger.info(f"ElationBot initialized for platform: {self.platform}")
def initialize(self):
"""Initialize the bot"""
try:
self.settings.validate_config()
self.settings.ensure_directories()
self.driver = self.web_driver_manager.get_driver()
self.wait = WebDriverWait(self.driver, 10)
self.logger.info("Bot initialized successfully")
return True
except Exception as e:
self.logger.error(f"Initialization failed: {str(e)}")
return False
def read_patient_data_from_excel(self, row_index=0):
"""Read patient information from OrderTemplate.xlsx for a specific row, including Document Id and PDF path"""
try:
# Read config to get OrderTemplate path
config_path = Path(__file__).parent.parent / 'config.json'
with open(config_path, 'r') as f:
config = json.load(f)
order_template_path = config['configuration']['OrderTemplatePath']
# Convert to absolute path from project root
project_root = Path(__file__).parent.parent
# If it's already a relative path, use it directly
if not order_template_path.startswith('/') and not (len(order_template_path) > 1 and order_template_path[1] == ':'):
order_template_path = project_root / order_template_path
else:
# Handle absolute paths by converting to relative
if 'elation-emr' in order_template_path:
# Split on 'elation-emr' and take the part after it
relative_part = order_template_path.split('elation-emr', 1)[1]
# Remove leading slash/backslash
relative_part = relative_part.lstrip('/\\')
order_template_path = project_root / relative_part
else:
# Fallback: try common relative paths
potential_paths = [
project_root / 'orders' / '2025-06-05' / 'docsathome_signed' / 'OrderTemplate.xlsx',
project_root / 'OrderTemplate.xlsx',
project_root / 'orders' / 'OrderTemplate.xlsx'
]
order_template_path = None
for path in potential_paths:
if path.exists():
order_template_path = path
break
if not order_template_path:
self.logger.error("Could not find OrderTemplate.xlsx in common locations")
return None
# Check if file exists
if not order_template_path.exists():
self.logger.error(f"OrderTemplate.xlsx not found at: {order_template_path}")
# Try alternative search in project directory
self.logger.info(f"Searching for OrderTemplate.xlsx in project directory: {project_root}")
# Search for the file recursively
xlsx_files = list(project_root.rglob('OrderTemplate.xlsx'))
if xlsx_files:
order_template_path = xlsx_files[0]
self.logger.info(f"Found OrderTemplate.xlsx at: {order_template_path}")
else:
self.logger.error("OrderTemplate.xlsx not found anywhere in the project directory")
return None
# Read Excel file
df = pd.read_excel(order_template_path)
self.logger.info(f"Successfully read OrderTemplate.xlsx with {len(df)} rows")
# Look for Patient Name and DOB columns (case insensitive)
patient_name_col = None
dob_col = None
for col in df.columns:
col_lower = str(col).lower()
if 'patient' in col_lower and 'name' in col_lower:
patient_name_col = col
elif 'dob' in col_lower or ('date' in col_lower and 'birth' in col_lower):
dob_col = col
if not patient_name_col:
self.logger.error("Could not find 'Patient Name' column in Excel file")
return None
if not dob_col:
self.logger.error("Could not find 'DOB' column in Excel file")
return None
# Check if row_index is valid
if len(df) == 0:
self.logger.error("No patient data found in Excel file")
return None
if row_index >= len(df):
self.logger.error(f"Row index {row_index} is out of range. Excel file has {len(df)} rows.")
return None
# Get the specified row of data
patient_data = df.iloc[row_index]
patient_name = str(patient_data[patient_name_col]).strip()
patient_dob = patient_data[dob_col]
# Format DOB if it's a datetime object
if pd.notna(patient_dob):
if hasattr(patient_dob, 'strftime'):
# If it's a datetime object, format it
patient_dob = patient_dob.strftime('%m/%d/%Y')
else:
# If it's already a string, clean it up
patient_dob = str(patient_dob).strip()
else:
patient_dob = None
# Get Document Id
document_id_col = None
for col in df.columns:
if 'document id' in str(col).lower():
document_id_col = col
break
if not document_id_col:
self.logger.error("Could not find 'Document Id' column in Excel file")
return None
document_id = str(patient_data[document_id_col]).strip()
project_root = Path(__file__).parent.parent
pdf_path = project_root / "files" / "SignedOrders" / f"{document_id}.pdf"
pdf_path = str(pdf_path.resolve())
self.logger.info(f"Read patient data (row {row_index}) - Name: {patient_name}, DOB: {patient_dob}, Document Id: {document_id}, PDF Path: {pdf_path}")
return {
'name': patient_name,
'dob': patient_dob,
'row_index': row_index,
'document_id': document_id,
'pdf_path': pdf_path
}
except Exception as e:
self.logger.error(f"Failed to read patient data from Excel: {str(e)}")
return None
def read_all_patients_from_excel(self, max_patients=None):
"""Read patient information from OrderTemplate.xlsx with optional limit, including Document Id and PDF path"""
try:
# Read config to get OrderTemplate path
config_path = Path(__file__).parent.parent / 'config.json'
with open(config_path, 'r') as f:
config = json.load(f)
order_template_path = config['configuration']['OrderTemplatePath']
# Convert to absolute path from project root
project_root = Path(__file__).parent.parent
# If it's already a relative path, use it directly
if not order_template_path.startswith('/') and not (len(order_template_path) > 1 and order_template_path[1] == ':'):
order_template_path = project_root / order_template_path
else:
# Handle absolute paths by converting to relative
if 'elation-emr' in order_template_path:
# Split on 'elation-emr' and take the part after it
relative_part = order_template_path.split('elation-emr', 1)[1]
# Remove leading slash/backslash
relative_part = relative_part.lstrip('/\\')
order_template_path = project_root / relative_part
else:
# Fallback: try common relative paths
potential_paths = [
project_root / 'orders' / '2025-06-05' / 'docsathome_signed' / 'OrderTemplate.xlsx',
project_root / 'OrderTemplate.xlsx',
project_root / 'orders' / 'OrderTemplate.xlsx'
]
order_template_path = None
for path in potential_paths:
if path.exists():
order_template_path = path
break
if not order_template_path:
self.logger.error("Could not find OrderTemplate.xlsx in common locations")
return []
# Check if file exists
if not order_template_path.exists():
self.logger.error(f"OrderTemplate.xlsx not found at: {order_template_path}")
# Try alternative search in project directory
self.logger.info(f"Searching for OrderTemplate.xlsx in project directory: {project_root}")
# Search for the file recursively
xlsx_files = list(project_root.rglob('OrderTemplate.xlsx'))
if xlsx_files:
order_template_path = xlsx_files[0]
self.logger.info(f"Found OrderTemplate.xlsx at: {order_template_path}")
else:
self.logger.error("OrderTemplate.xlsx not found anywhere in the project directory")
return []
# Read Excel file
df = pd.read_excel(order_template_path)
self.logger.info(f"Successfully read OrderTemplate.xlsx with {len(df)} rows")
# Look for Patient Name and DOB columns (case insensitive)
patient_name_col = None
dob_col = None
for col in df.columns:
col_lower = str(col).lower()
if 'patient' in col_lower and 'name' in col_lower:
patient_name_col = col
elif 'dob' in col_lower or ('date' in col_lower and 'birth' in col_lower):
dob_col = col
if not patient_name_col:
self.logger.error("Could not find 'Patient Name' column in Excel file")
return []
if not dob_col:
self.logger.error("Could not find 'DOB' column in Excel file")
return []
# Get all rows of data
if len(df) == 0:
self.logger.error("No patient data found in Excel file")
return []
patients = []
for index, row in df.iterrows():
try:
# Stop if we've reached the maximum number of patients
if max_patients is not None and len(patients) >= max_patients:
break
patient_name = str(row[patient_name_col]).strip()
patient_dob = row[dob_col]
# Skip rows with empty patient names
if not patient_name or patient_name.lower() in ['nan', 'none', '']:
continue
# Format DOB if it's a datetime object
if pd.notna(patient_dob):
if hasattr(patient_dob, 'strftime'):
# If it's a datetime object, format it
patient_dob = patient_dob.strftime('%m/%d/%Y')
else:
# If it's already a string, clean it up
patient_dob = str(patient_dob).strip()
else:
patient_dob = None
# Get Document Id
document_id_col = None
for col in df.columns:
if 'document id' in str(col).lower():
document_id_col = col
break
if not document_id_col:
self.logger.error("Could not find 'Document Id' column in Excel file")
continue
document_id = str(row[document_id_col]).strip()
project_root = Path(__file__).parent.parent
pdf_path = project_root / "files" / "SignedOrders" / f"{document_id}.pdf"
pdf_path = str(pdf_path.resolve())
patients.append({
'name': patient_name,
'dob': patient_dob,
'row_index': index,
'document_id': document_id,
'pdf_path': pdf_path
})
except Exception as row_e:
self.logger.warning(f"Error processing row {index}: {str(row_e)}")
continue
limit_msg = f" (limited to {max_patients})" if max_patients else ""
self.logger.info(f"Read {len(patients)} valid patients from Excel{limit_msg}")
return patients
except Exception as e:
self.logger.error(f"Failed to read all patients from Excel: {str(e)}")
return []
def _find_element(self, selectors, timeout=5):
"""Find element with multiple selectors"""
for selector in selectors:
try:
element = WebDriverWait(self.driver, timeout).until(
EC.element_to_be_clickable(selector)
)
if element:
return element
except:
continue
return None
def _find_search_bar(self, timeout=30):
"""Specifically find the search bar with detailed logging"""
search_selectors = [
(By.CSS_SELECTOR, '#chart-home-patient-search > div > span > span'), # User provided selector
(By.CSS_SELECTOR, '#chart-home-patient-search > div > span > span input'), # User selector + input
(By.CSS_SELECTOR, '#chart-home-patient-search input'), # Input within the specific container
(By.CSS_SELECTOR, '#chart-home-patient-search'), # The container
(By.CSS_SELECTOR, 'input[placeholder*="Find patient chart" i]'),
(By.CSS_SELECTOR, 'input[placeholder*="patient" i]'),
(By.CSS_SELECTOR, 'input[placeholder*="search" i]')
]
self.logger.info(f"Looking for search bar with {len(search_selectors)} selectors...")
for i, selector in enumerate(search_selectors):
try:
self.logger.debug(f"Trying selector {i+1}: {selector[1]}")
element = WebDriverWait(self.driver, 2).until(
EC.presence_of_element_located(selector)
)
if element and element.is_displayed():
self.logger.info(f"✅ Found search bar with selector {i+1}: {selector[1]}")
# If it's not an input, look for input inside
if element.tag_name != 'input':
try:
actual_input = element.find_element(By.CSS_SELECTOR, 'input')
self.logger.info("Found input field inside the container")
return actual_input
except:
# Try clicking to activate
element.click()
time.sleep(0.5)
try:
actual_input = element.find_element(By.CSS_SELECTOR, 'input')
self.logger.info("Found input field after clicking container")
return actual_input
except:
self.logger.debug("No input found inside container")
return element
else:
return element
except Exception as e:
self.logger.debug(f"Selector {i+1} failed: {str(e)}")
continue
self.logger.error("❌ Search bar not found with any selector")
return None
def login(self, username=None, password=None, url=None):
"""Login to Elation EMR with Google Authenticator support"""
try:
login_url = url or self.settings.ELATION_URL
login_username = username or self.settings.ELATION_USERNAME
login_password = password or self.settings.ELATION_PASSWORD
if not all([login_url, login_username, login_password]):
raise ValueError("Missing credentials")
self.logger.info("Starting login...")
self.driver.get(login_url)
# Username entry
username_selectors = [
(By.NAME, "username"),
(By.NAME, "email"),
(By.CSS_SELECTOR, 'input[type="email"]'),
(By.ID, "username")
]
username_element = self._find_element(username_selectors, timeout=8)
if not username_element:
self.logger.error("Username field not found")
return False
username_element.clear()
username_element.send_keys(login_username)
# Next button
next_selectors = [
(By.CSS_SELECTOR, 'input[type="submit"]'),
(By.CSS_SELECTOR, 'button[type="submit"]'),
(By.XPATH, "//button[contains(text(), 'Next')]")
]
next_button = self._find_element(next_selectors, timeout=3)
if next_button:
next_button.click()
else:
username_element.send_keys(Keys.RETURN)
time.sleep(1)
# Password entry
password_selectors = [
(By.NAME, "password"),
(By.CSS_SELECTOR, 'input[type="password"]'),
(By.ID, "password")
]
password_element = self._find_element(password_selectors, timeout=8)
if not password_element:
self.logger.error("Password field not found")
return False
password_element.clear()
password_element.send_keys(login_password)
# Click "Remember me" checkbox
try:
remember_me_selector = (By.CSS_SELECTOR, "#form66 > div.o-form-content.o-form-theme.clearfix > div.o-form-fieldset-container > div.o-form-fieldset.o-form-label-top.margin-btm-0 > div > span > div > label")
remember_me_element = WebDriverWait(self.driver, 5).until(
EC.element_to_be_clickable(remember_me_selector)
)
remember_me_element.click()
self.logger.info("✅ Clicked 'Remember me' checkbox")
except Exception as e:
self.logger.warning(f"Could not click 'Remember me' checkbox: {str(e)}")
# Continue with login even if remember me fails
# Login submission
login_selectors = [
(By.CSS_SELECTOR, 'input[type="submit"]'),
(By.CSS_SELECTOR, 'button[type="submit"]'),
(By.XPATH, "//button[contains(text(), 'Login')]")
]
login_button = self._find_element(login_selectors, timeout=3)
if login_button:
login_button.click()
else:
password_element.send_keys(Keys.RETURN)
# Initial verification to see if 2FA is required
time.sleep(3)
current_url = self.driver.current_url.lower()
# First check if we're already logged in successfully (search bar is visible)
try:
search_selectors = [
(By.CSS_SELECTOR, '#chart-home-patient-search input'), # Input within the specific container
(By.CSS_SELECTOR, '#chart-home-patient-search'), # The exact element you provided
(By.CSS_SELECTOR, 'input[placeholder*="Find patient chart" i]'),
(By.CSS_SELECTOR, 'input[placeholder*="patient" i]'),
(By.CSS_SELECTOR, 'input[placeholder*="search" i]')
]
# Try to find search bar with a short timeout
search_box = self._find_search_bar(timeout=3)
if search_box:
self.logger.info("✅ Login successful - search bar found!")
return True
except Exception as e:
self.logger.debug(f"Search bar not found yet: {str(e)}")
# Check if we're on a 2FA/authenticator page
auth_keywords = ['2fa', 'authenticator', 'verification', 'verify', 'code', 'token']
page_source = self.driver.page_source.lower()
needs_2fa = any(keyword in page_source for keyword in auth_keywords)
# Also check if we're still on login page or redirected to 2FA page
if 'login' in current_url or 'signin' in current_url or needs_2fa:
# Check specifically for Google Authenticator or 2FA elements
two_fa_selectors = [
(By.CSS_SELECTOR, 'input[placeholder*="code" i]'),
(By.CSS_SELECTOR, 'input[placeholder*="authenticator" i]'),
(By.CSS_SELECTOR, 'input[placeholder*="verification" i]'),
(By.NAME, "code"),
(By.NAME, "token"),
(By.NAME, "otp"),
(By.ID, "code"),
(By.ID, "token"),
(By.ID, "otp")
]
two_fa_element = None
for selector in two_fa_selectors:
try:
two_fa_element = WebDriverWait(self.driver, 2).until(
EC.presence_of_element_located(selector)
)
break
except:
continue
if two_fa_element:
self.logger.info("🔐 Google Authenticator required!")
print("\n" + "="*60)
print("🔐 GOOGLE AUTHENTICATOR REQUIRED")
print("="*60)
print("Please open your Google Authenticator app and")
print("enter the 6-digit code in the browser.")
print("\nWaiting 10 minutes for you to complete 2FA...")
print("The automation will continue after you submit the code.")
print("="*60)
# Wait for 10 minutes (600 seconds) for user to enter 2FA code
wait_time = 600 # 10 minutes
start_time = time.time()
while time.time() - start_time < wait_time:
try:
# Check if we've been redirected away from the 2FA page
current_url = self.driver.current_url.lower()
if 'login' not in current_url and 'signin' not in current_url and 'verify' not in current_url:
self.logger.info("✅ 2FA completed successfully!")
print("✅ 2FA completed! Continuing with automation...")
return True
# Also check for search bar during 2FA wait
try:
search_box = self._find_search_bar(timeout=2)
if search_box:
self.logger.info("✅ 2FA completed - search bar found!")
print("✅ 2FA completed! Continuing with automation...")
return True
except:
pass
# Wait 10 seconds before checking again
time.sleep(10)
remaining_time = wait_time - (time.time() - start_time)
if remaining_time > 0:
print(f"⏳ Waiting... {int(remaining_time//60)}:{int(remaining_time%60):02d} remaining")
except Exception as e:
self.logger.debug(f"Error during 2FA wait: {str(e)}")
time.sleep(10)
# Final check after 10 minutes
current_url = self.driver.current_url.lower()
if 'login' not in current_url and 'signin' not in current_url and 'verify' not in current_url:
self.logger.info("✅ Login successful after 2FA!")
return True
else:
# One more check for search bar
try:
search_box = self._find_search_bar(timeout=3)
if search_box:
self.logger.info("✅ Login successful - search bar found after timeout!")
return True
except:
pass
self.logger.error("❌ 2FA timeout - please try again")
print("❌ 2FA timeout - automation stopped")
return False
else:
# No 2FA detected but still on login page - check once more for search bar with longer wait
try:
self.logger.info("No 2FA detected, checking for search bar...")
search_box = self._find_search_bar(timeout=8)
if search_box:
self.logger.info("✅ Login successful - search bar found!")
return True
except:
pass
self.logger.error("❌ Login failed - still on login page")
return False
else:
# We're not on login page, do a final check for search bar to confirm success
try:
search_box = self._find_search_bar(timeout=5)
if search_box:
self.logger.info("✅ Login successful - search bar found!")
return True
except:
pass
# Direct login success without 2FA (fallback)
self.logger.info("✅ Login successful!")
return True
except Exception as e:
self.logger.error(f"Login error: {str(e)}")
return False
def search_patient(self, patient_name=None, patient_id=None, patient_dob=None):
"""Search for patient using provided name and DOB and capture Attributed To name"""
try:
# If no parameters provided, raise an error (no longer reading from Excel by default)
if not patient_name and not patient_id:
raise ValueError("Patient name or ID required")
# Create search term combining name and DOB
search_term = patient_name
if patient_dob:
search_term = f"{patient_name} {patient_dob}"
if not search_term:
raise ValueError("Patient name or ID required")
self.logger.info(f"Searching for: {search_term}")
# Find search box using dedicated method with detailed logging
search_box = self._find_search_bar(timeout=10)
if not search_box:
self.logger.error("Search box not found")
return False
# Search using name and DOB
search_box.clear()
search_box.send_keys(search_term)
time.sleep(1.5)
# Try to capture "Attributed To" name using the specific selector provided by user
try:
attributed_element = WebDriverWait(self.driver, 5).until(
EC.presence_of_element_located((By.CSS_SELECTOR, "#floating-ui-2 > div > div > li:nth-child(1) > a > span > div > div > div.BaseMindreader__description___2XPj0"))
)
if attributed_element and attributed_element.text.strip():
full_text = attributed_element.text.strip()
# Extract only the name after "Attributed to:"
if "Attributed to:" in full_text:
self.attributed_to_name = full_text.split("Attributed to:", 1)[1].strip()
elif "Attributed To:" in full_text:
self.attributed_to_name = full_text.split("Attributed To:", 1)[1].strip()
else:
# If no "Attributed to:" label, try to extract name from the end
# Format might be "DOB: date, Name" - take everything after the last comma
if "," in full_text:
self.attributed_to_name = full_text.split(",")[-1].strip()
else:
self.attributed_to_name = full_text
self.logger.info(f"Captured Attributed To using specific selector: {self.attributed_to_name}")
except:
self.logger.debug("Specific attributed selector not found, trying fallback methods")
# Wait for search results to appear
time.sleep(2)
# Try the user provided specific selector first
result_selectors = [
(By.CSS_SELECTOR, '#floating-ui-2 > div > div > li:nth-child(1) > a'), # User provided selector
(By.CSS_SELECTOR, '#floating-ui-2 > div > div > li:first-child > a'), # Alternative with first-child
(By.CSS_SELECTOR, '#floating-ui-2 li:first-child a'), # Simplified version
(By.XPATH, f"//span[contains(text(), '{patient_name.split()[-1]}')]"),
(By.XPATH, f"//span[contains(text(), '{patient_name}')]"),
(By.CSS_SELECTOR, '.patient-result'),
(By.CSS_SELECTOR, '.patient-item')
]
for i, selector in enumerate(result_selectors):
try:
self.logger.debug(f"Trying patient result selector {i+1}: {selector[1]}")
if i < 3: # For the specific selectors, try direct click
element = WebDriverWait(self.driver, 5).until(
EC.element_to_be_clickable(selector)
)
if element and element.is_displayed():
self.logger.info(f"✅ Found patient result with selector {i+1}: {selector[1]}")
# Click the patient result
self.logger.info("Clicking first patient result")
element.click()
time.sleep(2)
# Check for new tab
if len(self.driver.window_handles) > 1:
self.driver.switch_to.window(self.driver.window_handles[-1])
self.logger.info("Switched to patient chart")
return True
else: # For fallback selectors, use the old logic
results = self.driver.find_elements(*selector)
for result in results:
if result.is_displayed() and result.is_enabled():
# If we didn't capture attributed name with specific selector, try fallback methods
if not self.attributed_to_name:
try:
# Look for "Attributed To" or similar text near the patient result
parent_element = result.find_element(By.XPATH, "./..")
attributed_selectors = [
".//span[contains(text(), 'Attributed To') or contains(text(), 'Provider') or contains(text(), 'Physician')]",
".//div[contains(text(), 'Attributed To') or contains(text(), 'Provider') or contains(text(), 'Physician')]",
".//td[contains(text(), 'Attributed To') or contains(text(), 'Provider') or contains(text(), 'Physician')]"
]
for attr_selector in attributed_selectors:
try:
attributed_elements = parent_element.find_elements(By.XPATH, attr_selector)
for attr_elem in attributed_elements:
attr_text = attr_elem.text.strip()
if attr_text:
# Extract the name part after "Attributed To:" or similar
if ":" in attr_text:
self.attributed_to_name = attr_text.split(":", 1)[1].strip()
else:
# Try to find the next sibling or nearby element with the name
next_sibling = attr_elem.find_element(By.XPATH, "./following-sibling::*[1]")
if next_sibling and next_sibling.text.strip():
self.attributed_to_name = next_sibling.text.strip()
if self.attributed_to_name:
self.logger.info(f"Captured Attributed To using fallback: {self.attributed_to_name}")
break
except:
continue
if self.attributed_to_name:
break
# If we couldn't find "Attributed To" in the current result, look in the broader search results area
if not self.attributed_to_name:
try:
search_results_area = self.driver.find_element(By.CSS_SELECTOR, '.search-results, .patient-list, .results')
attributed_text_selectors = [
".//span[contains(text(), 'Attributed') or contains(text(), 'Provider')]",
".//div[contains(text(), 'Attributed') or contains(text(), 'Provider')]",
".//td[contains(text(), 'Attributed') or contains(text(), 'Provider')]"
]
for selector in attributed_text_selectors:
try:
elements = search_results_area.find_elements(By.XPATH, selector)
for elem in elements:
text = elem.text.strip()
if ":" in text:
potential_name = text.split(":", 1)[1].strip()
if potential_name and len(potential_name.split()) >= 2: # Likely a full name
self.attributed_to_name = potential_name
self.logger.info(f"Captured Attributed To from search area: {self.attributed_to_name}")
break
except:
continue
if self.attributed_to_name:
break
except:
pass
except Exception as e:
self.logger.debug(f"Could not capture Attributed To name using fallback: {str(e)}")
# Now click the patient result
self.logger.info("Clicking patient result")
result.click()
time.sleep(2)
# Check for new tab
if len(self.driver.window_handles) > 1:
self.driver.switch_to.window(self.driver.window_handles[-1])
self.logger.info("Switched to patient chart")
return True
except Exception as e:
self.logger.debug(f"Selector {i+1} failed: {str(e)}")
continue
self.logger.error("No patient results found")
return False
except Exception as e:
self.logger.error(f"Patient search failed: {str(e)}")
return False
def _get_documents_path(self):
"""Get documents path"""
if self.is_windows:
# Try OneDrive Documents first
onedrive_documents = Path.home() / 'OneDrive' / 'Documents'
if onedrive_documents.exists():
return onedrive_documents
# Fall back to regular Documents
return Path.home() / 'Documents'
return Path.home() / 'Documents'
def _open_file_explorer(self, path):
"""Open file explorer"""
try:
if self.is_windows:
# Use explorer.exe with proper path formatting
path_str = str(path).replace('/', '\\')
subprocess.run(['explorer', path_str], check=False)
elif self.is_macos:
subprocess.run(['open', str(path)], check=False)
else:
# Try common Linux file managers
for command in ['nautilus', 'dolphin', 'thunar', 'xdg-open']:
try:
subprocess.run([command, str(path)], check=False)
break
except FileNotFoundError:
continue
# Wait for window to open
time.sleep(2)
return True
except Exception as e:
self.logger.error(f"Failed to open file explorer: {str(e)}")
return False
def _find_file_using_computer_vision(self, file_name, template_image='file_thumbnail.png'):
"""Use computer vision to locate the file in Explorer"""
try:
self.logger.info(f"Looking for file: {file_name}")
# Take a screenshot
screenshot = pyautogui.screenshot()
screenshot_np = np.array(screenshot)
# Convert to grayscale for template matching
screenshot_gray = np.dot(screenshot_np[...,:3], [0.2989, 0.5870, 0.1140])
# Load the template image
try:
template_pil = Image.open(template_image)
template_np = np.array(template_pil)
if len(template_np.shape) == 3:
template_gray = np.dot(template_np[...,:3], [0.2989, 0.5870, 0.1140])
else:
template_gray = template_np
except Exception as e:
self.logger.error(f"Could not load template image: {template_image}, error: {e}")
return None
# Perform template matching using normalized cross-correlation
result = feature.match_template(screenshot_gray, template_gray)
# Find the best match
max_loc = np.unravel_index(np.argmax(result), result.shape)
max_val = result[max_loc]
# If good match found
if max_val > 0.7: # Confidence threshold
# Get center of match (note: match_template returns (y, x) format)
h, w = template_gray.shape[:2]
center_x = max_loc[1] + w//2
center_y = max_loc[0] + h//2
self.logger.info(f"Found file at coordinates: ({center_x}, {center_y})")
return (center_x, center_y)
self.logger.warning("Could not find file icon")
return None
except Exception as e:
self.logger.error(f"File detection failed: {str(e)}")
return None
def _fallback_file_detection(self, file_name):
"""Fallback method using OCR and grid search"""
try:
self.logger.info("🔄 Using fallback detection methods...")
# Try OCR detection first
file_coords = self._find_file_using_ocr(file_name)
if file_coords:
return file_coords
# If OCR fails, try intelligent grid search
file_coords = self._find_file_using_grid_search()
if file_coords:
return file_coords
# Last resort: use common positions
screen_size = pyautogui.size()
common_positions = [
(100, 150), (100, 220), (100, 290), # Left column
(180, 150), (180, 220), (180, 290), # Second column
(260, 150), (260, 220), (260, 290), # Third column
]
for pos in common_positions:
if pos[0] < screen_size[0] and pos[1] < screen_size[1]:
self.logger.info(f"Using fallback position: {pos}")
return pos
return None
except Exception as e:
self.logger.error(f"Fallback detection failed: {str(e)}")
return None
def _find_file_using_ocr(self, file_name):
"""Use OCR to find file name text"""
try:
import pytesseract
from PIL import Image
# Take screenshot
screenshot = pyautogui.screenshot()
# Get OCR data with bounding boxes
ocr_data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
# Look for filename matches
file_base = Path(file_name).stem
for i, text in enumerate(ocr_data['text']):
if not text.strip():
continue
text_clean = text.strip().lower()
# Check if this text matches our file
if (file_base.lower() in text_clean or
text_clean in file_base.lower() or
any(part in text_clean for part in file_base.split() if len(part) > 3)):
# Get text position
x = ocr_data['left'][i] + ocr_data['width'][i] // 2
y = ocr_data['top'][i]
# Adjust to click on icon above text (typical macOS/Windows layout)
icon_y = max(50, y - 30)
self.logger.info(f"OCR found '{text}' matching '{file_name}' at: ({x}, {icon_y})")
return (x, icon_y)
except ImportError:
self.logger.debug("pytesseract not available for OCR")
except Exception as e:
self.logger.debug(f"OCR detection failed: {str(e)}")
return None
def _find_file_using_grid_search(self):
"""Intelligent grid search for file icons"""
try:
screen_size = pyautogui.size()
# Define search area (typical Finder content area)
if self.is_macos:
# macOS Finder window layout
search_left = 50
search_top = 100
search_right = screen_size[0] - 50
search_bottom = screen_size[1] - 100
else:
# Windows/Linux file explorer
search_left = 50
search_top = 80
search_right = screen_size[0] - 50
search_bottom = screen_size[1] - 80
# Create grid of potential icon positions
icon_spacing_x = 80
icon_spacing_y = 80
positions = []
x = search_left + 40 # Start with some margin
while x < search_right:
y = search_top + 40
while y < search_bottom:
positions.append((x, y))
y += icon_spacing_y
x += icon_spacing_x
self.logger.info(f"Grid search: testing {len(positions)} positions")
# Return first valid position (this could be enhanced with validation)
if positions:
return positions[0]
except Exception as e:
self.logger.debug(f"Grid search failed: {str(e)}")
return None
def _click_upload_button_with_computer_vision(self):
"""Use advanced computer vision to find and click upload button"""
try:
self.logger.info("🖱️ Using enhanced computer vision to detect upload button...")
# Take a screenshot
screenshot = pyautogui.screenshot()
screenshot_np = np.array(screenshot)
# Convert to grayscale
screenshot_gray = np.dot(screenshot_np[...,:3], [0.2989, 0.5870, 0.1140])
# Load the upload button template image
template_path = 'upload_button.png'
if not Path(template_path).exists():
self.logger.error(f"Upload button template not found: {template_path}")
return False
try:
template_pil = Image.open(template_path)
template_np = np.array(template_pil)
if len(template_np.shape) == 3:
template_gray = np.dot(template_np[...,:3], [0.2989, 0.5870, 0.1140])
else:
template_gray = template_np
except Exception as e:
self.logger.error(f"Could not load upload button template: {template_path}, error: {e}")
return False
self.logger.info(f"📋 Template loaded ({template_gray.shape}), using multi-scale matching...")
# Try multiple scales for template matching
scales = [1.0, 0.9, 1.1, 0.8, 1.2, 0.7, 1.3, 0.6, 1.4]
best_match = None
best_confidence = 0
for scale in scales:
try:
# Resize template
if scale != 1.0:
new_height = int(template_gray.shape[0] * scale)
new_width = int(template_gray.shape[1] * scale)
if new_width <= 0 or new_height <= 0:
continue
scaled_template = transform.resize(
template_gray,
(new_height, new_width),
anti_aliasing=True,
preserve_range=True
).astype(template_gray.dtype)
else:
scaled_template = template_gray
# Skip if template is larger than screenshot
if (scaled_template.shape[0] > screenshot_gray.shape[0] or
scaled_template.shape[1] > screenshot_gray.shape[1]):
continue
# Perform template matching
result = feature.match_template(screenshot_gray, scaled_template)
# Find best match
max_loc = np.unravel_index(np.argmax(result), result.shape)
confidence = result[max_loc]
if confidence > best_confidence:
best_confidence = confidence
h, w = scaled_template.shape[:2]
best_match = {
'confidence': confidence,
'location': (max_loc[1], max_loc[0]), # Convert (y,x) to (x,y)
'size': (w, h),
'scale': scale
}
self.logger.debug(f"Scale {scale}: confidence {confidence:.3f}")
except Exception as e:
self.logger.debug(f"Error with scale {scale}: {str(e)}")
continue
# Check if we found a good match
confidence_threshold = 0.5 # Lower threshold for better detection
if best_match and best_match['confidence'] > confidence_threshold:
# Calculate center of the button
center_x = best_match['location'][0] + best_match['size'][0] // 2
center_y = best_match['location'][1] + best_match['size'][1] // 2
self.logger.info(f"✅ Upload button found at ({center_x}, {center_y}) with confidence {best_match['confidence']:.3f} (scale: {best_match['scale']})")
# Drag mouse to button and tap once
self.logger.info("🖱️ Moving mouse to upload button...")
pyautogui.moveTo(center_x, center_y, duration=1.0)
self.logger.info("👆 Tapping upload button once...")
pyautogui.click()
self.logger.info("✅ Upload button tapped")
return True
# If template matching fails, try OCR-based detection
self.logger.warning("Template matching failed, trying OCR detection...")
ocr_result = self._find_upload_button_using_ocr(screenshot)
if ocr_result:
center_x, center_y = ocr_result
self.logger.info(f"✅ Upload button found via OCR at ({center_x}, {center_y})")
pyautogui.moveTo(center_x, center_y, duration=1.0)
pyautogui.click()
self.logger.info("✅ Upload button tapped via OCR")
return True
# If OCR fails, try grid-based search
self.logger.warning("OCR failed, trying grid search...")
grid_result = self._find_upload_button_using_grid_search()
if grid_result:
center_x, center_y = grid_result
self.logger.info(f"✅ Upload button found via grid search at ({center_x}, {center_y})")
pyautogui.moveTo(center_x, center_y, duration=1.0)
pyautogui.click()
self.logger.info("✅ Upload button tapped via grid search")
return True
confidence = best_match['confidence'] if best_match else 0
self.logger.error(f"❌ Upload button not found with any method (best confidence: {confidence:.3f})")
return False
except Exception as e:
self.logger.error(f"Enhanced computer vision upload button detection failed: {str(e)}")
return False
def _find_upload_button_using_ocr(self, screenshot):
"""Use OCR to find upload button text"""
try:
import pytesseract
from PIL import Image
# Convert screenshot to PIL Image
screenshot_pil = Image.fromarray(np.array(screenshot))
# Get OCR data with bounding boxes
ocr_data = pytesseract.image_to_data(screenshot_pil, output_type=pytesseract.Output.DICT)
# Look for upload-related text
upload_keywords = ['upload', 'file', 'submit', 'save', 'attach']
for i, text in enumerate(ocr_data['text']):
if not text.strip():
continue
text_clean = text.strip().lower()
# Check if this text matches upload keywords
if any(keyword in text_clean for keyword in upload_keywords):
# Get text position
x = ocr_data['left'][i] + ocr_data['width'][i] // 2
y = ocr_data['top'][i] + ocr_data['height'][i] // 2
self.logger.info(f"OCR found '{text}' at: ({x}, {y})")
return (x, y)
except ImportError:
self.logger.debug("pytesseract not available for OCR")
except Exception as e:
self.logger.debug(f"OCR detection failed: {str(e)}")
return None
def _find_upload_button_using_grid_search(self):
"""Intelligent grid search for upload button in dialog area"""
try:
screen_size = pyautogui.size()
# Focus on dialog area (typically center of screen)
dialog_left = screen_size[0] // 4
dialog_top = screen_size[1] // 4
dialog_right = 3 * screen_size[0] // 4
dialog_bottom = 3 * screen_size[1] // 4
# Create grid of potential button positions in dialog area
button_spacing_x = 60
button_spacing_y = 40
positions = []
y = dialog_bottom - 100 # Start from bottom of dialog (where buttons usually are)
while y > dialog_top:
x = dialog_left + 50
while x < dialog_right:
positions.append((x, y))
x += button_spacing_x
y -= button_spacing_y
self.logger.info(f"Grid search: testing {len(positions)} positions in dialog area")
# Return first position as a fallback (bottom-right area of dialog)
if positions:
# Try the bottom-right area first (common button location)
priority_positions = [
(dialog_right - 100, dialog_bottom - 50),
(dialog_right - 200, dialog_bottom - 50),
(dialog_right - 150, dialog_bottom - 80)
]
for pos in priority_positions:
if (dialog_left < pos[0] < dialog_right and
dialog_top < pos[1] < dialog_bottom):
return pos
return positions[0]
except Exception as e:
self.logger.debug(f"Grid search failed: {str(e)}")
return None
def _drag_and_drop_file(self, file_path):
"""Smart drag and drop file upload"""
try:
self.logger.info(f"Starting drag and drop for: {file_path}")
# Get file info
file_name = Path(file_path).name
documents_path = self._get_documents_path()
documents_file_path = documents_path / file_name
# Copy file to documents if needed
if not documents_file_path.exists():
import shutil
shutil.copy2(file_path, documents_file_path)
self.logger.info(f"Copied file to documents: {documents_file_path}")
time.sleep(1)
# Open file explorer
if not self._open_file_explorer(documents_path):
return False
# Wait for Explorer to load
time.sleep(2)
# Alt-tab to file explorer to ensure it's focused
if self.is_windows:
pyautogui.hotkey('alt', 'tab')
time.sleep(1)
elif self.is_macos:
pyautogui.hotkey('command', 'tab')
time.sleep(1)
# Find file using computer vision
file_coords = self._find_file_using_computer_vision(file_name)
if not file_coords:
self.logger.error("Could not locate file")
return False
# Get screen size
screen_size = pyautogui.size()
# First click to select
pyautogui.moveTo(file_coords[0], file_coords[1])
time.sleep(0.5)
pyautogui.click()
time.sleep(0.5)
# Second click and hold to start drag
pyautogui.moveTo(file_coords[0], file_coords[1])
time.sleep(0.5)
pyautogui.mouseDown(button='left')
time.sleep(0.5)
# Small movement to initiate drag
pyautogui.moveRel(10, 0, duration=0.2)
time.sleep(0.5)
# Move to center of screen and drop
drop_x = screen_size[0] // 2
drop_y = screen_size[1] // 2
pyautogui.moveTo(drop_x, drop_y, duration=1.5)
time.sleep(0.5)
# Release to drop
pyautogui.mouseUp(button='left')
time.sleep(1)
self.logger.info("Drag and drop completed")
time.sleep(2)
# Return the temporary file path for later cleanup
return str(documents_file_path)
except Exception as e:
self.logger.error(f"Drag and drop failed: {str(e)}")
try:
pyautogui.mouseUp(button='left')
except:
pass
return False
def _handle_popup_form(self):
"""Handle the popup form that appears after drag and drop"""
try:
self.logger.info("Handling popup form after file upload...")
# Wait for popup to fully load
time.sleep(3) # Match previous method's timing
# Fill Provider field using the specific selector from previous method
if self.attributed_to_name:
try:
self.logger.info("📝 Filling Provider field...")
provider_container = WebDriverWait(self.driver, 10).until(
EC.presence_of_element_located((By.CSS_SELECTOR, "#ui-id-4 > div.dialog-content > div > form > div:nth-child(1) > div.el8FieldSection.ebs-form-group"))
)
provider_field = provider_container.find_element(By.CSS_SELECTOR, "input")
provider_field.clear()
provider_field.send_keys(self.attributed_to_name)
self.logger.info(f"Typed Provider name: {self.attributed_to_name}")
# Wait for autocomplete dropdown
time.sleep(1)
try:
first_suggestion = WebDriverWait(self.driver, 3).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "[id^='physicianUserName-popover-'] > div > div.mr-results > div:first-child > div"))
)
suggestion_text = first_suggestion.text.strip()
self.logger.info(f"Found autocomplete suggestion: '{suggestion_text}'")
first_suggestion.click()
self.logger.info(f"Selected autocomplete suggestion: '{suggestion_text}'")
except Exception as e:
self.logger.warning(f"Could not select autocomplete suggestion: {str(e)}")
self.logger.info("Trying keyboard navigation...")
try:
provider_field.send_keys(Keys.ARROW_DOWN)
time.sleep(1)
provider_field.send_keys(Keys.ENTER)
self.logger.info("Selected first option using keyboard navigation")
except Exception as ke:
self.logger.error(f"Keyboard navigation failed: {str(ke)}")
time.sleep(0.5)
except Exception as e:
self.logger.error(f"Failed to fill Provider field: {str(e)}")
# Select "Home Health Report" in Doc Type dropdown (using new method's logic)
try:
self.logger.info("📝 Selecting Doc Type dropdown...")
popup_container = WebDriverWait(self.driver, 10).until(
EC.presence_of_element_located((By.ID, "ui-id-4"))
)
doc_type_selectors = [
"select",
"select[name*='type']",
"select[name*='doc']",
".ebs-form-group select",
"#ui-id-4 select"
]
doc_type_dropdown = None
for selector in doc_type_selectors:
try:
doc_type_dropdown = popup_container.find_element(By.CSS_SELECTOR, selector)
break
except:
continue
if doc_type_dropdown:
self.driver.execute_script("arguments[0].scrollIntoView(true);", doc_type_dropdown)
time.sleep(0.5)
doc_type_dropdown.click()
self.logger.info("Opened Doc Type dropdown")
time.sleep(0.5)
select = Select(doc_type_dropdown)
try:
select.select_by_visible_text("Home Health Report")
self.logger.info("Successfully selected 'Home Health Report' in Doc Type dropdown")
except:
for option in select.options:
if "home health" in option.text.lower():
select.select_by_visible_text(option.text)
self.logger.info(f"Successfully selected '{option.text}' in Doc Type dropdown")
break
else:
self.logger.warning("Could not select 'Home Health Report' in dropdown")
time.sleep(0.5)
else:
self.logger.error("Doc Type dropdown not found in popup")
except Exception as e:
self.logger.error(f"Error handling Doc Type dropdown: {str(e)}")
# Fill Title field with document type using previous method's selector
try:
self.logger.info("📝 Filling Title field with document type...")
try:
from document_type_detector import main as get_document_type
document_info = get_document_type()
self.logger.info(f"Document type: {document_info}")
except Exception as e:
document_info = ''
self.logger.error(f"Document type not found: {str(e)}")
title_textarea = WebDriverWait(self.driver, 10).until(
EC.presence_of_element_located((By.CSS_SELECTOR, "#ui-id-4 > div.dialog-content > div > form > div:nth-child(6) > div > div > textarea.w100.ebs-form-control.info-field.el8InfoText"))
)
self.driver.execute_script("arguments[0].scrollIntoView(true);", title_textarea)
time.sleep(0.5)
title_textarea.clear()
title_textarea.send_keys(document_info or "Home Health Report")
self.logger.info(f"Successfully filled Title field with: {document_info or 'Home Health Report'}")
time.sleep(0.5)
except Exception as e:
self.logger.error(f"Failed to fill Title field: {str(e)}")
# Check "Mark on behalf of reviewer" checkbox using previous method's selector
try:
self.logger.info("☑️ Checking 'Mark on behalf of reviewer' checkbox...")
reviewer_label = WebDriverWait(self.driver, 5).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#ui-id-4 > div.dialog-content > div > form > div.el8FieldSection.ebs-form-group > div > ul > li:nth-child(2) > label"))
)
label_text = reviewer_label.text.lower()
if "reviewer" in label_text or "behalf of reviewer" in label_text:
self.driver.execute_script("arguments[0].scrollIntoView(true);", reviewer_label)
time.sleep(0.5)
reviewer_label.click()
self.logger.info("Successfully checked 'Mark on behalf of reviewer' checkbox")
try:
reviewer_checkbox_input = reviewer_label.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
if reviewer_checkbox_input.is_selected():
self.logger.info("✅ Reviewer checkbox confirmed as checked")
else:
self.logger.warning("⚠️ Reviewer checkbox not checked, retrying...")
reviewer_checkbox_input.click()
except:
self.logger.info("Could not verify checkbox state")
time.sleep(0.5)
else:
self.logger.warning(f"Found element but text doesn't match reviewer checkbox: '{label_text}'")
raise Exception("Element found but text doesn't match 'reviewer' checkbox")
except Exception as e:
self.logger.error(f"Failed to check reviewer checkbox: {str(e)}")
self.logger.info("🔴 Manual intervention required: giving 10 seconds...")
print("\n" + "="*60)
print("⚠️ Manual Action Required")
print("="*60)
print("Please manually click the 'Mark on behalf of reviewer' checkbox")
print("You have 10 seconds...")
print("="*60)
time.sleep(10)
print("\n✅ Resuming...")
# Click Upload button using previous method's selector
try:
self.logger.info("🎯 Clicking Upload button...")
upload_button = WebDriverWait(self.driver, 10).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#ui-id-4 > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > ul > li:nth-child(1) > button"))
)
self.driver.execute_script("arguments[0].scrollIntoView(true);", upload_button)
time.sleep(0.5)
upload_button.click()
self.logger.info("Successfully clicked Upload button")
time.sleep(4) # Wait for upload to process
except Exception as e:
self.logger.error(f"Failed to click Upload button: {str(e)}")
return False
# Refresh page and check chronological records using previous method's selector
self.logger.info("🔄 Refreshing page to check chronological records...")
self.driver.refresh()
time.sleep(5) # Match previous method's timing
try:
chart_feed_list = WebDriverWait(self.driver, 15).until(
EC.presence_of_element_located((By.CSS_SELECTOR, "#chart-feed-list"))
)
self.logger.info("Chart feed list found, checking for Home Health record...")
found_home_health = False
try:
chart_feed_text = chart_feed_list.text.lower()
if "home health" in chart_feed_text:
found_home_health = True
self.logger.info("Found Home Health record in chronological records!")
else:
feed_elements = chart_feed_list.find_elements(By.CSS_SELECTOR, "*")
for element in feed_elements:
try:
element_text = element.text.lower()
if "home health" in element_text:
found_home_health = True
self.logger.info("Found Home Health record in chronological records!")
break
except:
continue
if not found_home_health:
self.logger.warning(f"Home Health record not found in chronological records. Preview: {chart_feed_list.text[:200]}...")
except Exception as e:
self.logger.warning(f"Error checking chart feed text: {str(e)}")
except Exception as e:
self.logger.error(f"Chart feed list not found: {str(e)}")
found_home_health = False
# Switch back to main tab instead of navigating to homepage
self.logger.info("🔄 Switching back to main tab...")
self._switch_to_main_tab()
return found_home_health
except Exception as e:
self.logger.error(f"Error handling popup form: {str(e)}")
self._switch_to_main_tab() # Attempt to switch back even on error
return False
def _cleanup_temporary_file(self, temp_file_path):
"""Clean up temporary file after successful upload verification"""
try:
if temp_file_path and Path(temp_file_path).exists():
Path(temp_file_path).unlink()
self.logger.info(f"✅ Cleaned up temporary file: {temp_file_path}")
except Exception as e:
self.logger.warning(f"Failed to clean up temporary file {temp_file_path}: {str(e)}")
def _log_upload_to_csv(self, file_path, physician_name, document_type, title_field, remarks):
log_file = 'uploads_log.csv'
file_exists = Path(log_file).exists()
with open(log_file, mode='a', newline='', encoding='utf-8') as csvfile:
writer = csv.writer(csvfile)
if not file_exists:
writer.writerow(['File', 'Physician', 'Document Type', 'Title', 'Remarks'])
writer.writerow([file_path, physician_name, document_type, title_field, remarks])
# def upload_file(self, file_path):
# """Upload file using drag and drop and handle popup"""
# temp_file_path = None
# document_info = None
# remarks = ''
# try:
# if not Path(file_path).exists():
# remarks = f"File not found: {file_path}"
# self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', '', remarks)
# raise ValueError(remarks)
# self.logger.info(f"Uploading: {Path(file_path).name}")
# if len(self.driver.window_handles) > 1:
# self.driver.switch_to.window(self.driver.window_handles[-1])
# time.sleep(1)
# drag_result = self._drag_and_drop_file(file_path)
# if drag_result and drag_result != False:
# temp_file_path = drag_result
# self.logger.info("Drag and drop completed, handling popup...")
# # Try to get document info for logging
# try:
# from document_type_detector import main as get_document_type
# document_info = get_document_type()
# except Exception as e:
# document_info = ''
# verification_result = self._handle_popup_form()
# if verification_result:
# remarks = 'success'
# self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', document_info or '', remarks)
# self._cleanup_temporary_file(temp_file_path)
# return True
# else:
# remarks = 'chronological record not found after upload'
# self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', document_info or '', remarks)
# self.logger.warning("⚠️ Upload completed but verification failed - Home Health record not found")
# self.logger.info("📁 Keeping temporary file until manual verification")
# self.logger.info(f" Temporary file location: {temp_file_path}")
# return False
# else:
# remarks = 'upload failed (drag and drop failed)'
# self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', '', remarks)
# self.logger.error("Upload failed")
# return False
# except Exception as e:
# remarks = f"Upload error: {str(e)}"
# self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', document_info or '', remarks)
# self.logger.error(f"Upload error: {str(e)}")
# if temp_file_path:
# self._cleanup_temporary_file(temp_file_path)
# return False
# Replace the existing upload_file method with this method
def upload_file(self, file_path, is_batch=False):
"""Upload file using manual upload with delay and handle popup"""
temp_file_path = None
document_info = None
remarks = ''
try:
if not Path(file_path).exists():
remarks = f"File not found: {file_path}"
self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', '', remarks)
raise ValueError(remarks)
self.logger.info(f"Uploading: {Path(file_path).name}")
if len(self.driver.window_handles) > 1:
self.driver.switch_to.window(self.driver.window_handles[-1])
time.sleep(3)
# Manual upload prompt
print("\n" + "="*60)
print("📤 Manual File Upload Required")
print("="*60)
print(f"Please manually upload the file: {Path(file_path).name}")
if is_batch:
print("Note: In batch mode, prepare all files in advance to streamline uploads.")
print("1. Open the file explorer and locate the file.")
print("2. Drag and drop the file into the EHR upload area.")
print(f"You have {30 if not is_batch else 45} seconds to complete the upload.")
print("="*60)
self.logger.info(f"Pausing for {30 if not is_batch else 45} seconds to allow manual upload of {Path(file_path).name}")
time.sleep(30 if not is_batch else 45) # Longer delay in batch mode
self.logger.info("Resuming after manual upload delay")
print("✅ Resuming automation after manual upload...")
# Try to get document info for logging
try:
from document_type_detector import main as get_document_type
document_info = get_document_type()
time.sleep(3) # Brief pause after getting document info
except Exception as e:
document_info = ''
# Handle the popup form that appears after upload
self.logger.info("Handling popup form after manual file upload...")
time.sleep(5) # Wait for popup to appear
verification_result = self._handle_popup_form()
time.sleep(2) # Wait for popup processing to complete
if verification_result:
remarks = 'success'
self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', document_info or '', remarks)
self._cleanup_temporary_file(temp_file_path) # No temp file, but keep for compatibility
return True
else:
remarks = 'chronological record not found after upload'
self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', document_info or '', remarks)
self.logger.warning("⚠️ Upload completed but verification failed - Home Health record not found")
self.logger.info("📁 No temporary file created for manual upload")
return False
except Exception as e:
remarks = f"Upload error: {str(e)}"
self._log_upload_to_csv(file_path, self.attributed_to_name or '', 'Home Health Report', document_info or '', remarks)
self.logger.error(f"Upload error: {str(e)}")
return False
# Add this method to ElationBot class
def _reset_browser_state(self):
"""Reset browser to homepage and verify search bar is present"""
try:
self.logger.info("🔄 Resetting browser state to homepage...")
homepage_link = WebDriverWait(self.driver, 10).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#queuenav > a > div > span"))
)
homepage_link.click()
time.sleep(3) # Wait for homepage to load
search_box = self._find_search_bar(timeout=5)
if search_box:
self.logger.info("✅ Browser state reset: homepage loaded, search bar found")
return True
else:
self.logger.warning("⚠️ Search bar not found after resetting to homepage")
return False
except Exception as e:
self.logger.error(f"Failed to reset browser state: {str(e)}")
return False
# Add this method to ElationBot class
def _switch_to_main_tab(self):
"""Close the current tab and switch back to the main tab"""
try:
self.logger.info("🔄 Closing current tab and switching to main tab...")
if len(self.driver.window_handles) > 1:
current_handle = self.driver.current_window_handle
self.driver.close() # Close the patient chart tab
self.driver.switch_to.window(self.driver.window_handles[0]) # Switch to main tab
time.sleep(1) # Wait for tab switch to stabilize
# Verify search bar is present
search_box = self._find_search_bar(timeout=5)
if search_box:
self.logger.info("✅ Switched to main tab, search bar found")
return True
else:
self.logger.warning("⚠️ Search bar not found in main tab")
return False
else:
self.logger.warning("⚠️ Only one tab open, cannot switch")
return False
except Exception as e:
self.logger.error(f"Failed to switch to main tab: {str(e)}")
return False
# Replace the existing run_workflow method with this
def run_workflow(self, file_path=None, patient_name=None, patient_dob=None, username=None, password=None, url=None, keep_open=True, skip_login=False):
"""Run complete workflow - reads patient data from Excel if not provided"""
try:
self.logger.info("Starting workflow...")
# If patient data not provided, read from Excel
if not patient_name:
self.logger.info("Patient data not provided, reading from Excel...")
patient_data = self.read_patient_data_from_excel()
if not patient_data:
raise ValueError("Could not read patient data from Excel file")
patient_name = patient_data.get('name')
patient_dob = patient_data.get('dob')
file_path = patient_data.get('pdf_path')
if not patient_name:
raise ValueError("Patient name not found in Excel file")
self.logger.info(f"Read from Excel - Patient: {patient_name}, DOB: {patient_dob or 'Not provided'}, PDF: {file_path}")
else:
# If file_path is not provided, try to get it from Excel for the given patient
if not file_path:
self.logger.info("File path not provided, trying to get from Excel for given patient...")
patient_data = self.read_patient_data_from_excel()
if patient_data:
file_path = patient_data.get('pdf_path')
file_path = str(Path(file_path).resolve())
# Initialize the bot (skip if already initialized)
if not self.driver and not self.initialize():
if not keep_open:
self.close()
return False
# Skip login if requested and already logged in
if not skip_login and not self.login(username, password, url):
if not keep_open:
self.close()
return False
# Search for patient using data from Excel or provided parameters
self.logger.info(f"Searching for patient: {patient_name} (DOB: {patient_dob or 'Not provided'})")
if not self.search_patient(patient_name=patient_name, patient_dob=patient_dob):
return False
# Upload file
self.logger.info(f"Uploading: {Path(file_path).name}")
if not self.upload_file(file_path, is_batch=(not keep_open)):
return False
# Handle keep_open parameter - only close on success if keep_open
if not keep_open:
self.close()
print("Workflow completed successfully. Bot closed.")
return True
except Exception as e:
self.logger.error(f"Workflow failed: {str(e)}")
return False
def run_batch_workflow(self, file_paths=None, username=None, password=None, url=None, keep_open=False, max_patients=2):
"""Run batch workflow for multiple patients and files from Excel, using correct PDF path"""
try:
self.logger.info("🚀 Starting batch workflow...")
# Read patients from Excel with limit
self.logger.info(f"📊 Reading patients from Excel (max {max_patients})...")
patients = self.read_all_patients_from_excel(max_patients=max_patients)
if not patients:
raise ValueError("No patients found in Excel file")
# Use the correct PDF path for each patient
files_to_process = [p['pdf_path'] for p in patients]
# Validate that we have files for all patients
for file_path in files_to_process:
if not Path(file_path).exists():
self.logger.error(f"File not found: {file_path}")
return False
self.logger.info(f"📋 Processing {len(patients)} patients with {len(files_to_process)} files (max limit: {max_patients})")
successful_uploads = 0
failed_uploads = 0
for i, (patient, file_path) in enumerate(zip(patients, files_to_process)):
try:
self.logger.info(f"\n{'='*60}")
self.logger.info(f"📋 PROCESSING PATIENT {i+1}/{len(patients)}")
self.logger.info(f"👤 Patient: {patient['name']}")
self.logger.info(f"📅 DOB: {patient['dob'] or 'Not provided'}")
self.logger.info(f"📁 File: {Path(file_path).name}")
self.logger.info(f"{'='*60}")
if not self.search_patient(patient_name=patient['name'], patient_dob=patient['dob']):
self.logger.error(f"❌ Failed to find patient: {patient['name']}")
failed_uploads += 1
continue
self.logger.info(f"📤 Uploading file: {Path(file_path).name}")
if self.upload_file(file_path):
self.logger.info(f"✅ Successfully uploaded file for {patient['name']}")
successful_uploads += 1
else:
self.logger.error(f"❌ Failed to upload file for {patient['name']}")
failed_uploads += 1
except Exception as patient_e:
self.logger.error(f"❌ Error processing patient {patient['name']}: {str(patient_e)}")
failed_uploads += 1
continue
return successful_uploads > 0
except Exception as e:
self.logger.error(f"Batch workflow failed: {str(e)}")
if not keep_open:
self.close()
return False
def close(self):
"""Clean shutdown"""
if self.web_driver_manager:
self.web_driver_manager.quit()
self.logger.info("Bot closed")