from datetime import datetime
from datetime import timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
import json
import os
import time

STATE_FILE = "processed_dates.json"
load_dotenv()
driver_path = os.getenv("CHROME_DRIVER_PATH")
secret_password = os.getenv('PASSWORD')

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True) # Selenium don't close the Chrome browser when the script finishes
# options.headless = True  # Enable headless mode

service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)
url = os.getenv('SOSS_URL')
driver.get(url)
wait = WebDriverWait(driver, 15)

"""start logging in"""

try:
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'signInName'))
    )
except:
    print("Nie znaleziono elementu w czasie oczekiwania.")

time.sleep(2)
driver.find_element(By.CSS_SELECTOR, ".cky-btn.cky-btn-accept").click()
driver.find_element(By.XPATH, '//*[@id="signInName"]').send_keys('karol.grabiec@ergohestia.pl')
driver.find_element(By.ID, 'continue').click()
wait.until(EC.presence_of_element_located((By.ID, "i0116"))).send_keys('karol.grabiec@ergohestia.pl')

try:
    button = wait.until(EC.presence_of_element_located((By.ID, "idSIButton9")))
    driver.execute_script("arguments[0].click();", button)
except TimeoutException:
    print("Przycisk 'Tak' nie pojawił się w zadanym czasie.")

wait.until(EC.presence_of_element_located((By.ID, "passwordInput"))).send_keys(secret_password)
driver.find_element(By.ID, 'submitButton').click()
wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

"""end of login"""

wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Wyszukaj sprawę")]'))).click()
time.sleep(3)
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ocsg-dynamic-filters__input"))).click()
option_xpath = "//div[contains(@class, 'ant-select-item-option-content') and normalize-space()='Stan sprawy']"

try:
    option = wait.until(EC.visibility_of_element_located((By.XPATH, option_xpath)))
    option.click()
except:
    option = driver.find_element(By.XPATH, option_xpath)
    driver.execute_script("arguments[0].click();", option)

status_select = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, "nz-select-top-control.ant-select-selector")))
status_select.click()
status_option_xpath = "//div[contains(@class, 'ant-select-item-option-content') and normalize-space()='Zamknięta']"

try:
    option = wait.until(EC.visibility_of_element_located((By.XPATH, status_option_xpath)))
    option.click()
except:
    option = driver.find_element(By.XPATH, status_option_xpath)
    driver.execute_script("arguments[0].click();", option)

element_confirm = driver.find_element(By.XPATH, "//span[normalize-space()='Zastosuj']")
driver.execute_script("arguments[0].click();", element_confirm)
filter_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Dodaj filtr']")))
filter_trigger.click()
option_xpath = "//div[contains(@class, 'ant-select-item-option-content') and normalize-space()='Data wpływu do EH']"

try:
    option = wait.until(EC.visibility_of_element_located((By.XPATH, option_xpath)))
    option.click()
except:
    option = driver.find_element(By.XPATH, option_xpath)
    driver.execute_script("arguments[0].click();", option)

def get_date(message):
    while True:
        try:
            data_str = input(message)
            # Validation of date format and existence (e.g. whether January 32nd is not present)
            data_obj = datetime.strptime(data_str, '%Y-%m-%d')
            # length check
            if len(data_str) != 10:
                print("Błąd: Data musi mieć dokładnie 10 znaków")
                continue

            return data_obj  # return the datetime object if everything is fine
        except ValueError:
            print(f"Nieprawidłowy format lub data: {data_str}. Spróbuj ponownie.")

# calculating all days in a range
start_datetime = get_date("Podaj datę początkową wpływu (rrrr-mm-dd): ")
end_datetime = get_date("Podaj datę końcową wpływu (rrrr-mm-dd): ")

while start_datetime <= end_datetime:
    data_str = start_datetime.strftime('%Y-%m-%d')
    # load received dates
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, 'r') as f:
            processed_dates = set(json.load(f))
    else:
        processed_dates = set()

    if data_str not in processed_dates:
        received_date_to = wait.until(EC.element_to_be_clickable((By.XPATH, "(//input[@id='data-wplywu-do-eh-1-ctrl'])[1]")))
        received_date_to.send_keys(data_str)
        received_date_to = wait.until(EC.element_to_be_clickable((By.XPATH, "(//input[@id='data-wplywu-do-eh-1-ctrl'])[2]")))
        received_date_to.send_keys(data_str)
        element_confirm = wait.until(EC.element_to_be_clickable((By.ID, "submit-inflowDate")))
        driver.execute_script("arguments[0].click();", element_confirm)
        driver.find_element(By.XPATH, "//span[normalize-space()='Szukaj']").click()

    else:
        print(f"Data {data_str} już była przetworzona, pomijam...")
        start_datetime = start_datetime + timedelta(days=1)
        continue

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ocsg-dynamic-filters__input"))).click()
    filter_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Dodaj filtr']")))
    filter_trigger.click()
    option_xpath_case_status = "//div[contains(@class, 'ant-select-item-option-content') and normalize-space()='Status sprawy']"

    try:
        option = wait.until(EC.visibility_of_element_located((By.XPATH, option_xpath_case_status)))
        option.click()
    except:
        option = driver.find_element(By.XPATH, option_xpath_case_status)
        driver.execute_script("arguments[0].click();", option)

    wait.until(EC.element_to_be_clickable((By.ID, "status-sprawy-1-ctrl"))).click()
    status_ids = {
        "Realizacja wyroku - po I instancji": "status-sprawy-1-SENTENCE_REALIZATION_FIRST_INSTANCE",
        "Zakończona - bez wnoszenia apelacj": "status-sprawy-1-CLOSED_WITHOUT_APPEAL",
        "Realizacja wyroku - po II instancji": "status-sprawy-1-SENTENCE_REALIZATION_SECOND_INSTANCE"
    }
    for k, _id in status_ids.items():
        try:
            wait.until(EC.visibility_of_element_located((By.ID, _id))).click()
            print(f"Zaznaczono: {k}")
        except Exception as e:
            print(f"Problem z {k}, próbuję przez JS...")
            target = driver.find_element(By.ID, _id)
            driver.execute_script("arguments[0].click();", target)
    button_confirm = wait.until(EC.element_to_be_clickable((By.ID, "submit-status")))
    driver.execute_script("arguments[0].click();", button_confirm)

    case_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='/view']")
    case_urls = [el.get_attribute("href") for el in case_elements]
    print(f"Znaleziono spraw: {len(case_urls)}")

    # check and write received dates
    if data_str not in processed_dates:
        processed_dates.add(data_str)
        with open(STATE_FILE, 'w') as f:
            json.dump(list(processed_dates), f)

    start_datetime = start_datetime + timedelta(days=1)
    print(start_datetime)

for url in case_urls:
    print(f"Wchodzę do sprawy: {url}")
    driver.get(url)
    SELECTOR_DAMAGE = "ocsg-output[label*='DAMAGE_NUMBER'] p"
    try:
        case_damage_number_el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, SELECTOR_DAMAGE))).text.strip()
        print(f"Pobrano numer/numery spraw: {case_damage_number_el}")
    except:
        case_damage_number_el = "Nie znaleziono"

    signature_selector = "ocsg-output[label='CCM.CASE.COURT_FILE_REFERENCE'] p"
    try:
        signature_el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, signature_selector))).text.strip()
        print(f"Pobrano sygnaturę: {signature_el}")
    except:
        signature_el = "Nie znaleziono"

    try:
        xpath_grup = "//div[contains(text(), 'Grupa ubezpieczeń')]/following-sibling::div//p"

        grupa_el = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_grup)))
        grup_text = grupa_el.text.strip()

        # if the text is empty (for example it is in the tooltip), get the innerText via JS
        if not grup_text:
            grup_text = driver.execute_script("return arguments[0].innerText;", grupa_el).strip()
        print(f"🔎 Odczytana Grupa Ubezpieczeń: {grup_text}")

    except Exception as e:
        print(f"⚠️ Nie udało się odczytać pola Grupa ubezpieczeń: {e}")

    rows = driver.find_elements(By.CSS_SELECTOR, "app-case-documents table tbody tr")

    for row in rows:
        try:
            # get all cells in a row and skip those that don't have enough cells
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) < 2:
                continue

            # get the category from the second cell
            category = cells[2].text.strip()

            if "z uzasadnieniem" in category.lower():
                print(f"✅ Znaleziono: {category}")
                expand_icon = cells[0].find_element(By.CSS_SELECTOR, "button.ant-table-row-expand-icon")
                driver.execute_script("arguments[0].click();", expand_icon)
                print("➕ Rozwinięto szczegóły wiersza.")
                download_icon = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "svg[data-icon='download']")))
                driver.execute_script(
                    "arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles: true, cancelable: true}));",
                    download_icon) #send a 'click' event instead of calling the .click() function
                print("💾 Ikona pobierania kliknięta!")
                time.sleep(5)

        except Exception as e:
            print(f"⚠️ Nie udało się pobrać dokumentu: {e}")
            continue


    time.sleep(5)
time.sleep(10)