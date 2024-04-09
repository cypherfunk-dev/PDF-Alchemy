
# my_pdf = "/Users/cypher.funk/Documents/Git/PDFto/frontend/image-based-pdf-sample.pdf"

# doc = fitz.open(my_pdf) 
# def pdftype(doc):
#     i=0
#     for page in doc:
#         if len(page.getText())>0: #for scanned page it will be 0
#             i+=1
#     if i>0:
#         print('full/partial text PDF file')
#     else:
#         print('only scanned images in PDF file')
# pdftype(doc)