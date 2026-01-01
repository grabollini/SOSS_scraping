import json
import logging
import os
import sys
import time
from datetime import datetime
from datetime import timedelta

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from PIL import Image
from docx2pdf import convert

STATE_FILE = "processed_dates.json"
load_dotenv()
DRIVER_PATH = os.getenv("CHROME_DRIVER_PATH")
SECRET_PASSWORD = os.getenv('PASSWORD')
SECRET_LOGIN = os.getenv("LOGIN_MAIL")
USER_DATA = os.getenv('USER_DATA_DIR')
DOWNLOAD_DIR = os.getenv('NEW_DOWNLOAD_DIR')
all_cases = []
site_url = os.getenv("SHAREPOINT_LIST_URL")
list_name = "BAZA WYROKÓW PDF"

os.system("taskkill /F /IM chrome.exe /T >nul 2>&1")
os.system("taskkill /F /IM chromedriver.exe /T >nul 2>&1")


def setup_driver():
    print("\n" + "=" * 30)
    print("🤖 SOSS BOT - WYBÓR TRYBU")
    print("=" * 30)
    print("1. Nowa instancja (Kopia profilu)")
    print("2. Podłączenie do istniejącej instancji (Port 9222)")
    print("=" * 30)

    choice = input("Wybierz opcję (1/2): ")

    chrome_options = Options()
    service = Service(executable_path=DRIVER_PATH)  # Używamy Twojej stałej ścieżki

    if choice == '1':
        print("\n🚀 Uruchamiam nową sesję z Twoimi ustawieniami...")

        # Twoje czyszczenie procesów przed startem
        os.system("taskkill /F /IM chrome.exe /T >nul 2>&1")
        os.system("taskkill /F /IM chromedriver.exe /T >nul 2>&1")

        # Twoje preferencje pobierania
        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.exit_type": "Normal",
            "profile.exited_cleanly": True
        }

        # Twoja pełna lista flag
        chrome_options.add_argument(f"--user-data-dir={USER_DATA}")
        chrome_options.add_argument("--profile-directory=Default")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--remote-debugging-port=9222")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-session-crashed-bubble")

        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_experimental_option("prefs", prefs)

    elif choice == '2':
        print("\n🔗 Próbuję przejąć kontrolę nad otwartym Chrome...")
        # W tym trybie łączymy się tylko po adresie debuggera
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

    else:
        print("❌ Nieprawidłowy wybór. Zamykam.")
        sys.exit()

    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"\n❌ BŁĄD STARTU: {e}")
        if choice == '2':
            print("\nUpewnij się, że Chrome został uruchomiony z flagą --remote-debugging-port=9222")
        sys.exit()

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.exit_type": "Normal",
    "profile.exited_cleanly": True
}

chrome_options = Options()
chrome_options.add_argument(f"--user-data-dir={USER_DATA}")
chrome_options.add_argument("--profile-directory=Default")
chrome_options.add_argument("--disable-blink-features=AutomationControlled") # removes the navigator.webdriver flag from the bot so that the server does not recognize the bot
# these 3 flags often solve the DevToolsActivePort problem:
chrome_options.add_argument("--remote-debugging-port=9222") # opens a port for communication
chrome_options.add_argument("--no-sandbox")                # disables sandboxing (often required for profiles)
chrome_options.add_argument("--disable-dev-shm-usage")     # revents shared memory problems
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True) # Selenium don't close the Chrome browser when the script finishes
chrome_options.add_argument("--disable-session-crashed-bubble")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])# disabling the toolbar "Chrome is controlled by automatic software"
chrome_options.add_experimental_option('useAutomationExtension', False) # blocks the automation extension
# chrome_options.add_argument("--headless=new")  # Enable headless mode
chrome_options.add_experimental_option("prefs", prefs)

service = Service(DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)
url = os.getenv('SOSS_URL')
driver.get(url)
wait = WebDriverWait(driver, 5)
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") # cleans up bot traces in JS after starting the driver

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


def process_file_to_pdf(file_path):  # Converts file to PDF depending on extension.
    ext = os.path.splitext(file_path)[1].lower()
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"

    if ext == ".pdf":
        return file_path

    try:
        if ext in [".tiff", ".tif"]:
            print(f"🖼️ Konwertuję TIFF na PDF...")
            img = Image.open(file_path)
            if img.mode != 'RGB':  # Conversion to RGB - required for PDF and multi-page support
                img = img.convert('RGB')
            img.save(pdf_path, "PDF", resolution=300.0, save_all=True)
            os.remove(file_path)  # Delete the original
            return pdf_path

        elif ext == ".docx":
            print(f"📄 Konwertuję Word na PDF...")
            convert(file_path, pdf_path)
            os.remove(file_path)  # Remove the original
            return pdf_path
    except Exception as e:
        print(f"❌ Błąd konwersji pliku {ext}: {e}")
        return file_path

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
            print(f"Zaznaczono {k} przez JS")

    button_confirm = wait.until(EC.element_to_be_clickable((By.ID, "submit-status")))
    driver.execute_script("arguments[0].click();", button_confirm)
    time.sleep(5)
    try:
        search_btn_xpath = "//button[.//span[contains(text(), 'Szukaj')]]"
        # wait.until(EC.element_to_be_clickable((By.ID, "submit-status"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, search_btn_xpath))).click()

    except Exception as e:
        print(f"⚠️ Problem z {e}, klikam przez JS...")
        try:
            search_element = driver.find_element(By.XPATH, "//button[.//span[contains(text(), 'Szukaj')]]")
            driver.execute_script("arguments[0].click();", search_element)
            print("⚡ Przycisk 'Szukaj' kliknięty przez JS.")
        except Exception as final_e:
            print(f"💀 Nie udało się kliknąć przycisku: {final_e}")

    time.sleep(5)
    case_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='/view']")
    raw_urls = [el.get_attribute("href") for el in case_elements]
    case_urls = [url for url in set(raw_urls) if url and "http" in url]
    print(f"📊 Raport wyszukiwania:")
    print(f"   - Wszystkie znalezione elementy: {len(case_elements)}")
    print(f"   - Elementy z poprawnym linkiem: {len(case_urls)}")

    if not case_urls:
        print("⚠️ Uwaga: Nie znaleziono żadnych linków do spraw! Sprawdź czy strona się załadowała.")
    time.sleep(3)

    case_urls = [el.get_attribute("href") for el in case_elements if el.get_attribute("href")]
    print(f"Znaleziono spraw: {len(case_urls)}")

    time.sleep(3)
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
        case_id = wait.until(EC.visibility_of_element_located((By.ID, "mainHeadingId"))).text.strip()
        print(f"🆔 ID sprawy: {case_id}")
    except Exception as e:
        print(f"❌ Nie udało się pobrać ID: {e}")

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

    currency = wps_value.split()[-1]
    print(f"💵 Waluta: {currency}")

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
                global files_before
                files_before = set(os.listdir(DOWNLOAD_DIR))
                expand_icon = cells[0].find_element(By.CSS_SELECTOR, "button.ant-table-row-expand-icon")
                driver.execute_script("arguments[0].click();", expand_icon)
                print("➕ Rozwinięto szczegóły wiersza.")
                download_icon = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "svg[data-icon='download']")))
                driver.execute_script(
                    "arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles: true, cancelable: true}));",
                    download_icon) # send a 'click' event instead of calling the .click() function
                print("💾 Ikona pobierania kliknięta!")
                time.sleep(10)

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

    try:
        new_file_path = None
        for _ in range(15):
            time.sleep(1)
            files_after = set(os.listdir(DOWNLOAD_DIR))
            new_files = files_after - files_before
            valid_files = [f for f in new_files if not f.endswith(('.tmp', '.crdownload'))]  # file filtration

            if valid_files:
                original_name = valid_files[0]
                old_path = os.path.join(DOWNLOAD_DIR, original_name)

                # Changing the name to a unique one (keeping the original extension)
                extension = os.path.splitext(original_name)[1]
                temp_case_path = os.path.join(DOWNLOAD_DIR, f"Case_{case_id}{extension}")
                os.rename(old_path, temp_case_path)

                # Converting to PDF
                final_pdf_path = process_file_to_pdf(temp_case_path)
                print(f"✅ Plik gotowy: {os.path.basename(final_pdf_path)}")
                break

    except Exception as e:
        print(f"⚠️ Błąd w procesie pliku: {e}")

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

    case_data = {
        "Title": case_id,
        "StatusSprawy": status_text,
        "NumerSprawy": case_damage_number_el,
        "Sygnatura": signature_el,
        "WPS": wps_value,
        "Waluta": currency,
        "GrupaUbezpieczen": grupa_ubezpieczen,
        "RodzajRoszczenia": rodzaj_roszczenia,
        "KwotaZasadzona": kwota_zasadzona,
        "Instancja": court_verdict
    }
    all_cases.append(case_data)

    for v in all_cases:
        print(json.dumps(v, ensure_ascii=False))

    print(f"\n{'=' * 20} KONIEC SPRAWY {'=' * 20}\n")


"""
    ctx = ClientContext(site_url).with_credentials(UserCredential(SECRET_LOGIN, SECRET_PASSWORD))
    def send_to_sharepoint(data_dict):
        try:
            target_list = ctx.web.lists.get_by_title(list_name)
            item_create = target_list.add_item(data_dict).execute_query() # creating a new row
            print(f"✅ Dane wysłane do SharePoint: {data_dict.get('Title')}")
        except Exception as e:
            print(f"❌ Błąd SharePoint: {e}")
"""
time.sleep(5)
time.sleep(10)