import win32com.client
import pymupdf as fitz
import easyocr
import os
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

EASYOCR_READER = None
reader = easyocr.Reader(['pl', 'en'], gpu=False)

# load arial font
try:
    pdfmetrics.registerFont(TTFont('Arial-Pol', 'Arial.ttf'))
    pdfmetrics.registerFont(TTFont('Arial-Bold-Pol', 'Arialbd.ttf'))
    FONT_NAME = 'Arial-Pol'
    FONT_NAME_BOLD = 'Arial-Bold-Pol'
except:
    # Backup, if the script is running on Linux or another system
    FONT_NAME = 'Helvetica'
    FONT_NAME_BOLD = 'Helvetica-Bold'
    print("⚠️ Nie znaleziono czcionki Arial, używam domyślnej (brak PL znaków)")


def extract_attachments_from_msg(msg_path, download_dir):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(os.path.abspath(msg_path))
    extracted_files = []

    for attachment in msg.Attachments:
        file_name = attachment.FileName
        if file_name.lower().endswith(('.pdf', '.tif', '.tiff', '.fax', '.docx', '.doc')):
            save_path = os.path.join(download_dir, file_name)
            attachment.SaveAsFile(save_path)
            extracted_files.append(save_path)
            print(f"📎 Wyciągnięto załącznik: {file_name}")

    return extracted_files


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
                    is_dirty = False

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
                            is_dirty = True

                        full_pdf_text += text + "\n"
                        # Saves the text to the "output_text" folder for your viewing.
                        # with open(f"output_text/{file_name}.txt", "w", encoding="utf-8") as f:
                        #     f.write(full_pdf_text)
                if is_dirty:
                    new_page = doc.new_page()
                    new_page.insert_text((50, 50), full_pdf_text, fontsize=8, color=(1, 1, 1))
                    doc.saveIncr()
                    print(f"💾 Warstwa tekstowa została dopisana do pliku: {file_name}")
                doc.close()

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


def inject_metadata_to_pdf(source_pdf_path, metadata_json):
    """
    Create a new first page using JSON data and insert it into the beginner PDF file.
    """
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)

    # Font and header settings
    can.setFont(FONT_NAME_BOLD, 16)
    can.drawString(50, 800, "KARTA ANALITYCZNA WYROKU")
    can.line(50, 790, 550, 790)

    # Input data from JSON
    can.setFont(FONT_NAME, 12)
    y_position = 760

    # Mapping data from JSON
    for key, value in metadata_json.items():
        text = f"{key.replace('_', ' ').upper()}: {value}"
        can.drawString(50, y_position, text)
        y_position -= 20  # New line

        if y_position < 50:  # End of page protect
            can.showPage()
            can.setFont(FONT_NAME, 11)
            y_position = 800

    can.save()
    packet.seek(0)

    temp_path = source_pdf_path.replace(".pdf", "_final.pdf")

    try:
        # Use the 'with' block to ensure that files are properly closed
        with open(source_pdf_path, "rb") as existing_file:
            new_pdf = PdfReader(packet)
            existing_pdf = PdfReader(existing_file)
            output = PdfWriter()

            # First add a new metadata page.
            output.add_page(new_pdf.pages[0])

            # Add the rest of the pages of the original judgment (after OCR).
            for page in existing_pdf.pages:
                output.add_page(page)

            # Overwrite the file or create a new one
            with open(temp_path, "wb") as outputStream:
                output.write(outputStream)

        # File source_pdf_path is already closed – safe to delete.
        if os.path.exists(source_pdf_path):
            os.remove(source_pdf_path)
            print(f"🗑️ Usunięto plik przejściowy: {os.path.basename(source_pdf_path)}")

    except Exception as e:
        print(f"❌ Błąd podczas łączenia PDF i usuwania starego pliku: {e}")
        # W razie błędu zwracamy chociaż ścieżkę do tego, co udało się przetworzyć
        return source_pdf_path

    return temp_path


def process_file_to_pdf(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"

    if ext == ".pdf":
        return file_path

    try:
        # Image format group (TIFF, JPG, PNG)
        if ext in [".tiff", ".tif", ".jpg", ".jpeg", ".png", ".fax"]:
            print(f"🖼️ Konwertuję obraz {ext} na PDF...")
            pages = []
            with Image.open(file_path) as img:
                for p in range(getattr(img, 'n_frames', 1)):
                    img.seek(p)
                    page = img.convert('RGB')
                    pages.append(page)
                if pages:
                    pages[0].save(pdf_path, "PDF", resolution = 300.0, save_all = True, append_images = pages[1:])

                    # clear the list of objects in memory
                    for p in pages:
                        p.close()

            try:
                os.remove(file_path)
            except OSError:
                time.sleep(0.5) # Windows can block file saving when a file is sent to the cloud, so it is a good method to add a 0.5-second sleep.
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


def get_reader():
    """Initializes the EasyOCR reader only if needed."""
    global EASYOCR_READER
    if EASYOCR_READER is None:
        print("⏳ Inicjalizacja EasyOCR (pobieranie modeli językowych przy pierwszym uruchomieniu)...")
        EASYOCR_READER = easyocr.Reader(['pl', 'en'], gpu=False) # gpu=False is safer if you don't have a graphics card
    return EASYOCR_READER