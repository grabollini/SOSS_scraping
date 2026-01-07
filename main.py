import easyocr
import faiss
import pymupdf as fitz
import json
import numpy as np
import os
import time
from datetime import datetime
from datetime import timedelta

import comtypes.client
from PIL import Image
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from sentence_transformers import SentenceTransformer, util

STATE_FILE = "processed_dates.json"
load_dotenv()
DRIVER_PATH = os.getenv("CHROME_DRIVER_PATH")
SECRET_PASSWORD = os.getenv('PASSWORD')
SECRET_LOGIN = os.getenv("LOGIN_MAIL")
USER_DATA = os.getenv('USER_DATA_DIR')
DOWNLOAD_DIR = os.getenv('NEW_DOWNLOAD_DIR')
site_url = os.getenv("SHAREPOINT_LIST_URL")
list_name = "BAZA WYROKÓW PDF"
files_before = set(os.listdir(DOWNLOAD_DIR))
MODEL_PATH = r'C:\Model\all-MiniLM-L6-v2'
INDEX_FILE = 'faiss_index.bin'
MAPPING_FILE = 'faiss_mapping_ids.npy'
EASYOCR_READER = None
reader = easyocr.Reader(['pl', 'en'], gpu=False)

GLOBAL_MODEL = None
GLOBAL_FAISS_INDEX = None
GLOBAL_MAPPING_IDS = None
all_cases = []
case_urls = []

# Setting up offline mode for supporting libraries. SentenceTransformer model downloaded manually.
os.environ['TRANSFORMERS_OFFLINE'] = '1'
os.environ['HF_DATASETS_OFFLINE'] = '1'

os.system("taskkill /F /IM chrome.exe /T >nul 2>&1")
os.system("taskkill /F /IM chromedriver.exe /T >nul 2>&1")

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
chrome_options.add_argument("--headless=new")  # Enable headless mode
chrome_options.add_experimental_option("prefs", prefs)

service = Service(DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)
url = os.getenv('SOSS_URL')
driver.get(url)
wait = WebDriverWait(driver, 5)
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") # cleans up bot traces in JS after starting the driver


def clean_download_directory(directory_path):
    """
    Usuwa całą zawartość wskazanego folderu (pliki i podfoldery).
    Sam folder zostaje zachowany.
    """
    if not os.path.exists(directory_path):
        print(f"📂 Folder {directory_path} nie istnieje. Tworzę go...")
        os.makedirs(directory_path)
        return

    print(f"🧹 Czyszczenie folderu: {directory_path}")

    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        try:
            # Sprawdź czy to plik lub link symboliczny i usuń
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            # Jeśli to folder, usuń go wraz z zawartością
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"⚠️ Nie udało się usunąć {file_path}. Powód: {e}")

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
            print(f"Nieprawidłowy format lub data, spróbuj ponownie.")

# calculating all days in a range
start_datetime = get_date("❗ Podaj datę początkową wpływu do EH (rrrr-mm-dd): ")
end_datetime = get_date("❗ Podaj datę końcową wpływu do EH (rrrr-mm-dd): ")


def process_file_to_pdf(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"

    if ext == ".pdf":
        return file_path

    try:
        # Image format group (TIFF, JPG, PNG)
        if ext in [".tiff", ".tif", ".jpg", ".jpeg", ".png"]:
            print(f"🖼️ Konwertuję obraz {ext} na PDF...")

            with Image.open(file_path) as img:
                # --- DOWNLOAD VERIFICATION ---
                img.verify()  # Whether it is damage
                img = Image.open(file_path) # Open again to close file
                # ----------------------------

                if img.mode != 'RGB':
                    img = img.convert('RGB')

                img.save(pdf_path, "PDF", resolution=300.0, save_all=True) # save_all=True is important for multi-page TIFF

            os.remove(file_path)
            return pdf_path

        elif ext in [".docx", ".doc"]:
            time.sleep(5)
            print(f"📄 Konwertuję Word na PDF...")
            print(f"Ścieżka pliku Word: {file_path}")
            print('Jestem przed konwersją Word')
            abs_file_path = os.path.abspath(file_path)
            abs_pdf_path = os.path.abspath(pdf_path)
            word = None
            doc = None
            try:
                word = comtypes.client.CreateObject('Word.Application') # Run Word in the background
                word.Visible = False
                doc = word.Documents.Open(abs_file_path)
                doc.SaveAs(abs_pdf_path, FileFormat=17)
                print('✅ Konwersja zakończona sukcesem')

            except Exception as e:
                print(f"❌ Błąd konwersji Word: {e}")
                raise e
            finally:
                if doc:
                    doc.Close()
                if word:
                    word.Quit()

            os.remove(file_path)
            return pdf_path

    except Exception as e:
        print(f"❌ Błąd konwersji pliku {ext}: {e}")
        return file_path


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


while start_datetime <= end_datetime:
    clean_download_directory(DOWNLOAD_DIR)
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

    time.sleep(10)
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
        case_id = wait.until(EC.visibility_of_element_located((By.ID, "mainHeadingId"))).text.strip().replace("Sprawa ", "")
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
                expand_icon = cells[0].find_element(By.CSS_SELECTOR, "button.ant-table-row-expand-icon")
                driver.execute_script("arguments[0].click();", expand_icon)
                print("➕ Rozwinięto szczegóły wiersza.")
                download_icon = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "svg[data-icon='download']")))
                driver.execute_script(
                    "arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles: true, cancelable: true}));",
                    download_icon) # send a 'click' event instead of calling the .click() function
                print("💾 Ikona pobierania kliknięta!")
                time.sleep(3)

        except Exception as e:
            print(f"⚠️Nie udało się pobrać dokumentu: {e}")
            continue

    try:
        new_file_path = None
        for _ in range(3):
            time.sleep(1)
            print('Jestem w for _ in range')
            files_after = set(os.listdir(DOWNLOAD_DIR))
            print(f'After {files_after}')
            print(f'Before {files_before}')
            new_files = files_after - files_before
            valid_files = [f for f in new_files if not f.endswith(('.tmp', '.crdownload'))]  # file filtration

            if valid_files:
                print(f"✅ Znaleziono nowe pliki: {valid_files}")
                for original_filename in valid_files:
                    original_name_base = os.path.splitext(original_filename)[0]
                    original_name_base = original_name_base.replace(" ", "_").replace(".", "_")
                    old_path = os.path.join(DOWNLOAD_DIR, original_filename)

                    # Changing the name to a unique one (keeping the original extension)
                    extension = os.path.splitext(original_filename)[1]
                    safe_case_id = case_id.replace("/", "_").replace("\\", "_").replace(":", "_").replace(" ", "_")
                    temp_case_path = os.path.join(DOWNLOAD_DIR, safe_case_id + original_name_base + extension)
                    os.rename(old_path, temp_case_path)

                    # Converting to PDF
                    print('Jestem przed wywołaniem funkcji')
                    new_file_path = process_file_to_pdf(temp_case_path)
                    print('Jestem przed dodaniem do files_before')
                    files_before.add(os.path.basename(new_file_path))
                    print('Jestem po dodaniu do files_before')
                    print(f"✅ Plik po zmainie nazwy gotowy: {os.path.basename(new_file_path)}")
                    print(f"Nowa nazwa: {temp_case_path}")
                break
        time.sleep(5)
    except Exception as e:
        print(f"⚠️ Błąd w procesie pliku: {e}")

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

""" MINI LM"""
# print('Wyłącz sieć')
# time.sleep(20)
# Forcing offline mode
print("--- [SynLex] Rozpoczynanie ładowania komponentów offline... ---")

try:
    # Model loading
    if os.path.exists(MODEL_PATH):
        GLOBAL_MODEL = SentenceTransformer(MODEL_PATH, device='cpu') # Path to the folder, we avoid SSL/503 errors
        print("✅ Model semantyczny załadowany z dysku.")
    else:
        raise FileNotFoundError(f"Nie znaleziono folderu modelu w: {MODEL_PATH}")

    # Loading faiss index
    if os.path.exists(INDEX_FILE):
        GLOBAL_FAISS_INDEX = faiss.read_index(INDEX_FILE)
        print("✅ Indeks FAISS załadowany.")
    else:
        print("⚠️ Ostrzeżenie: Brak pliku indeksu .bin.")

    # 3. Loading mapping (file names)
    if os.path.exists(MAPPING_FILE):
        GLOBAL_MAPPING_IDS = np.load(MAPPING_FILE)
        print("✅ Mapowanie plików załadowane.")
    else:
        print("⚠️ Ostrzeżenie: Brak pliku mapowania .npy.")

    print("🚀 [SynLex] System gotowy do pracy.")

except Exception as e: # This block will catch an error if, for example, the MODEL_PATH folder does not exist or the
                        # files inside are corrupted.
    print(f"❌ [SynLex] BŁĄD KRYTYCZNY podczas startu: {e}")
    exit()


def get_reader():
    """Initializes the EasyOCR reader only if needed."""
    global EASYOCR_READER
    if EASYOCR_READER is None:
        print("⏳ Inicjalizacja EasyOCR (pobieranie modeli językowych przy pierwszym uruchomieniu)...")
        EASYOCR_READER = easyocr.Reader(['pl', 'en'], gpu=False) # gpu=False is safer if you don't have a graphics card
    return EASYOCR_READER


def create_semantic_index(pdf_folder):
    global GLOBAL_MODEL

    documents_text = []
    document_ids = []

    print(f"📄 Przeszukiwanie folderu: {pdf_folder}")

    for file_name in os.listdir(pdf_folder):
        if file_name.endswith('.pdf'):
            path = os.path.join(pdf_folder, file_name)
            try:
                full_pdf_text = ""
                with fitz.open(path) as doc:
                    for page in doc:
                        # Normal reading attempt
                        text = page.get_text().strip()

                        # Use EasyOCR if page is blank (scan)
                        if len(text) < 50:
                            print(f"🔍 Plik {file_name} (str. {page.number + 1}) to skan. Uruchamiam EasyOCR...")

                            # Rendering to an in-memory image
                            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                            img_bytes = pix.tobytes("png")

                            # Reading text via EasyOCR
                            reader = get_reader()
                            ocr_results = reader.readtext(img_bytes, detail=0)
                            text = " ".join(ocr_results)

                        full_pdf_text += text + "\n"
                        # Saves the text to the "output_text" folder for your viewing.
                        # with open(f"output_text/{file_name}.txt", "w", encoding="utf-8") as f:
                        #     f.write(full_pdf_text)


                if full_pdf_text.strip():
                    documents_text.append(full_pdf_text)
                    document_ids.append(file_name)
                    print(f"✅ Przetworzono: {file_name}")
                else:
                    print(f"⚠️ Pominięto: {file_name} (całkowicie pusty)")

            except Exception as e:
                print(f"❌ Błąd przy {file_name}: {e}")

    if not documents_text:
        print("⚠ Nie znaleziono żadnego tekstu. Indeks nie zostanie stworzony.")
        return

    # Embedding and FAISS
    print(f"🧠 Generowanie wektorów dla {len(documents_text)} dokumentów...")
    embeddings = GLOBAL_MODEL.encode(documents_text, show_progress_bar=True)
    embeddings = np.array(embeddings).astype('float32')

    dimension = embeddings.shape[1]
    index = faiss.IndexFlatL2(dimension)
    index.add(embeddings)

    # Saving (renaming files to your variables)
    faiss.write_index(index, 'faiss_index.bin')
    np.save('faiss_mapping_ids.npy', np.array(document_ids))
    print(f"🚀 Gotowe! Baza zawiera {index.ntotal} dokumentów.")

create_semantic_index(DOWNLOAD_DIR)
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