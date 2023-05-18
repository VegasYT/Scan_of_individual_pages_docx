import docx
import docx2txt
import docx2pdf
import os
import PyPDF2
import zipfile
from PIL import Image
from io import BytesIO
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

doc_path = 'document.docx' # укажите путь к документу здесь
pdf_path = 'document.pdf' # путь для сохранения PDF-файла

# Конвертируем документ в PDF
docx2pdf.convert(doc_path, pdf_path)

# Открываем PDF-файл и считываем первую страницу
with open(pdf_path, 'rb') as pdf_file:
    pdf_reader = PyPDF2.PdfFileReader(pdf_file)
    if pdf_reader.getNumPages() >= 2:
        num_pages = 2
    else:
        num_pages = 1
    for i in range(num_pages):
        page = pdf_reader.getPage(i)
        print(page.extractText())

# Удаляем PDF-файл
os.remove(pdf_path)


# Работа с картинками
doc = docx.Document(doc_path)
zip_file = zipfile.ZipFile(doc_path)
image_paths = [x for x in zip_file.namelist() if x.endswith(('jpeg', 'png', 'jpg', 'bmp'))]

for i in range(num_pages):
    page = pdf_reader.getPage(i)
    print(page.extractText())

for image_path in image_paths:
    image_data = zip_file.read(image_path)
    img = Image.open(BytesIO(image_data))
    byte_arr = BytesIO()
    img.save(byte_arr, format='PNG')
    bytes_img = byte_arr.getvalue()
    print(pytesseract.image_to_string(Image.open(BytesIO(bytes_img)), lang="rus"))

zip_file.close()
