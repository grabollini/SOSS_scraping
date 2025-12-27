import json
import os
import time
from datetime import datetime
from datetime import timedelta

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

STATE_FILE = "processed_dates.json"
load_dotenv()
DRIVER_PATH = os.getenv("CHROME_DRIVER_PATH")
SECRET_PASSWORD = os.getenv('PASSWORD')
USER_DATA = os.getenv('USER_DATA_DIR')

os.system("taskkill /F /IM chrome.exe /T >nul 2>&1")
os.system("taskkill /F /IM chromedriver.exe /T >nul 2>&1")

chrome_options = Options()
chrome_options.add_argument(f"--user-data-dir={USER_DATA}")
chrome_options.add_argument("--profile-directory=Default")
# these 3 flags often solve the DevToolsActivePort problem:
chrome_options.add_argument("--remote-debugging-port=9222") # opens a port for communication
chrome_options.add_argument("--no-sandbox")                # disables sandboxing (often required for profiles)
chrome_options.add_argument("--disable-dev-shm-usage")     # rrevents shared memory problems
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True) # Selenium don't close the Chrome browser when the script finishes
chrome_options.add_argument("--disable-session-crashed-bubble")
# chrome_options.add_argument("--headless=new")  # Enable headless mode

service = Service(DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)
url = os.getenv('SOSS_URL')
driver.get(url)
wait = WebDriverWait(driver, 15)

"""START OF LOGINING"""

try:
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'signInName'))
    )
except:
    print("Nie znaleziono elementu w czasie oczekiwania.")

time.sleep(2)
# driver.find_element(By.CSS_SELECTOR, ".cky-btn.cky-btn-accept").click()
driver.find_element(By.XPATH, '//*[@id="signInName"]').send_keys('karol.grabiec@ergohestia.pl')
driver.find_element(By.ID, 'continue').click()
try:
    login_input = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "i0116")))

    if login_field:
        print("🔐 Bot nie jest zalogowany. Rozpoczynam procedurę logowania...")
        login_input.send_keys('karol.grabiec@ergohestia.pl')
        try:
            button = wait.until(EC.presence_of_element_located((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", button)
        except TimeoutException:
            print("Przycisk 'Tak' nie pojawił się w zadanym czasie.")

        wait.until(EC.presence_of_element_located((By.ID, "passwordInput"))).send_keys(SECRET_PASSWORD)
        driver.find_element(By.ID, 'submitButton').click()
        wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    else:
        print("✅ Sesja aktywna. Pomijam logowanie.")

except Exception as e:
    print("✅ Zalogowano (brak pól logowania).")

"""END OF LOGINING"""

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
start_datetime = get_date("❗ Podaj datę początkową wpływu do EH (rrrr-mm-dd): ")
end_datetime = get_date("❗ Podaj datę końcową wpływu do EH (rrrr-mm-dd): ")

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

    """scraping data from the case cards with specific case status"""

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ocsg-dynamic-filters__input"))).click()
    filter_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Dodaj filtr']")))
    filter_trigger.click()

    try:
        option_xpath_case_status = "//div[contains(@class, 'ant-select-item-option-content') and normalize-space()='Status sprawy']"
        option = wait.until(EC.visibility_of_element_located((By.XPATH, option_xpath_case_status)))
        option.click()
    except:
        option = driver.find_element(By.XPATH, option_xpath_case_status)
        driver.execute_script("arguments[0].click();", option)

    wait.until(EC.element_to_be_clickable((By.ID, "status-sprawy-1-ctrl"))).click()
    status_ids = {
        "Realizacja wyroku - po I instancji": "status-sprawy-1-SENTENCE_REALIZATION_FIRST_INSTANCE",
        "Realizacja wyroku - po II instancji": "status-sprawy-1-SENTENCE_REALIZATION_SECOND_INSTANCE",
        "Zakończona - oddalenie powództwa": "status-sprawy-1-CLOSED_CLAIM_DISMISSED",
        "Zakończona - bez wnoszenia apelacji": "status-sprawy-1-CLOSED_WITHOUT_APPEAL",
        "Zakończona - brak ochrony": "status-sprawy-1-CLOSED_NO_COVERAGE"
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
    time.sleep(20)
    # check and write received dates
    if data_str not in processed_dates:
        processed_dates.add(data_str)
        with open(STATE_FILE, 'w') as f:
            json.dump(list(processed_dates), f)
    print(f"\n📆 Sprawdzana data: {start_datetime}\n")
    start_datetime = start_datetime + timedelta(days=1)

for url in case_urls:
    print(f"Wchodzę do sprawy: {url}")
    driver.get(url)

    try:
        status_element = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "nz-tag.ant-tag-error")))
        status_text = status_element.text.strip()
        status_text = status_text.replace("Status: ", "")
        print(f"📋 Status sprawy: {status_text}")

    except Exception as e:
        print(f"❌ Nie udało się pobrać statusu: {e}")

    try:
        selector_damage = "ocsg-output[label*='DAMAGE_NUMBER'] p"
        case_damage_number_el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, selector_damage))).text.strip()
        print(f"🆔 Numer/numery spraw: {case_damage_number_el}")
    except:
        case_damage_number_el = "Nie znaleziono"

    try:
        wps_selector = "ocsg-output[label='CCM.CASE.TOTAL_CASE_VALUE'] .ocsg-output__value"
        wps_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, wps_selector)))
        wps_value = wps_element.text.strip()
        print(f"💰 WPS (łączny): {wps_value}")
    except Exception as e:
        print(f"⚠️Nie udało się pobrać wartości WPS: {e}")

    try:
        signature_selector = "ocsg-output[label='CCM.CASE.COURT_FILE_REFERENCE'] p"
        signature_el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, signature_selector))).text.strip()
        print(f"📜 Sygnatura akt: {signature_el}")
    except:
        signature_el = "Nie znaleziono"

    try:
        table_xpath = "//table[thead//th[contains(text(), 'Grupa ubezpieczeń')]]"
        row_xpath = f"{table_xpath}/tbody/tr[1]"
        row = wait.until(EC.visibility_of_element_located((By.XPATH, row_xpath)))
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) >= 4:
            grupa_ubezpieczen = cols[0].text.strip()
            rodzaj_roszczenia = cols[1].text.strip()
            kwota_zasadzona = cols[3].text.strip()
            print(f"✅ Grupa ubezpieczeń: {grupa_ubezpieczen}")
            print(f"✅ Rodzaj roszczenia: {rodzaj_roszczenia}")
            print(f"✅ Kwota zasądzona: {kwota_zasadzona}")
        elif len(cols) >= 2: # security if the table has only 2-3 columns
            print(f"✅ Grupa ubezpieczeń: {cols[0].text.strip()}")
            print(f"✅ Rodzaj roszczenia: {cols[1].text.strip()}")
            print("ℹ️Brak kwoty zasądzenia.")
        else:
            print("⚠️Tabela znaleziona, ale jest zbyt krótka.")

    except Exception as e:
        print(f"❌ Błąd podczas pobierania danych z tabeli: {e}")

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
            print(f"⚠️Nie udało się pobrać dokumentu: {e}")
            continue

    try:
        court_1_inst = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "ocsg-output[label*='COURT_OF_FIRST_INSTANCE'] .ocsg-output__value")
        )).text.strip()
        court_2_inst = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "ocsg-output[label*='COURT_OF_SECOND_INSTANCE'] .ocsg-output__value")
        )).text.strip()

        print(f"🏛️ Sąd I instancji: {court_1_inst}")
        print(f"🏛️ Sąd II instancji: {court_2_inst}")

    except Exception as e:
        print(f"⚠️Nie udało się pobrać danych o sądach: {e}")

    if court_2_inst is None and category == "Wyrok I instancji z uzasadnieniem":
        court_verdict = court_1_inst
    elif court_2_inst is not None and category == "Wyrok II instancji z uzasadnieniem":
        court_verdict = court_2_inst
    else:
        court_verdict = court_1_inst
    print(f"⚖️Wyrok wydany przez: {court_verdict}")

    if kwota_zasadzona == "-" or kwota_zasadzona == "ℹ️Brak kwoty zasądzenia.":
        kwota_zasadzona = 0
    else:
        kwota_zasadzona = kwota_zasadzona.replace(" PLN", "").replace(" ", "").replace(",", ".")
        kwota_zasadzona = float(kwota_zasadzona)

    kwota_wps_clean = wps_value.replace(" PLN", "").replace(" ", "").replace(",", ".")
    wps_value = float(kwota_wps_clean)
    if status_text in ["Zakończona - oddalenie powództwa", "Zakończona - brak ochrony"]:
        wynik_wyroku = "Oddalenie"
    elif status_text in ["Realizacja wyroku - po I instancji", "Realizacja wyroku - po II instancji", "Zakończona - bez wnoszenia apelacji"] and wps_value > kwota_zasadzona:
        wynik_wyroku = "Uwzględnienie w części"
    elif status_text in ["Realizacja wyroku - po I instancji", "Realizacja wyroku - po II instancji", "Zakończona - bez wnoszenia apelacji"] and wps_value <= kwota_zasadzona:
        wynik_wyroku = "Uwzględnienie w całości"
    else:
        wynik_wyroku = "Brak danych do określenia wyniku wyroku"
    print(f"🏆 Wynik wyroku: {wynik_wyroku}")

    print(f"\n{'=' * 20} KONIEC SPRAWY {'=' * 20}\n")
    time.sleep(2)
time.sleep(10)