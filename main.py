from pdf2docx import Converter

pdf_path = "Add your PDF file"
docx_path = "Output Microsoft Word File"

cv = Converter(pdf_file=pdf_path)
cv.convert(docx_filename=docx_path)
cv.close()