import docx2pdf
import PyPDF2
import os

doc_path = 'document.docx' # укажите путь к документу здесь
pdf_path = 'document.pdf' # путь для сохранения PDF-файла

# Конвертируем документ в PDF
docx2pdf.convert(doc_path, pdf_path)

# Открываем PDF-файл и считываем первые две страницы
with open(pdf_path, 'rb') as pdf_file:
    pdf_reader = PyPDF2.PdfFileReader(pdf_file)
    for i in range(2):
        page = pdf_reader.getPage(i)
        print(page.extractText())

# Удаляем PDF-файл
os.remove(pdf_path)