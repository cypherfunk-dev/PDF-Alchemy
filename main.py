import customtkinter
from CTkMessagebox import CTkMessagebox
import tkinter
import os
import threading
import textwrap
from pdf2docx import Converter
import pandas as pd
from PIL import Image, ImageTk
import camelot
import time
import pandas as pd
import fitz # Assuming PyMuPDF is already installed
import ocrmypdf
import pytesseract

# Estilo e inicialización
pytesseract.pytesseract.tesseract_cmd=r"C:/Program Files/Tesseract-OCR/tesseract.exe"

customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("blue")

app = customtkinter.CTk()
app.geometry("600x500")
app.title('PDF Alchemy - Convert PDF to...')
app.resizable(False, False)

# Definición de variables globales
rutaarchivo = ""
botonruta = None
global imagen2

# Definición de funciones
def seleccionar_documento():
    global rutaarchivo
    global botonruta
    filename = customtkinter.filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")], )
    print(filename)
    rutaarchivo = os.path.splitext(filename)[0]
    if rutaarchivo == "":
        mensaje = etiquetaruta.configure(text="No files selected", font=("helvetica", 15))
        return mensaje
    botonruta.pack_forget()
    nombre = rutaarchivo + os.path.splitext(filename)[1]
    nombre_rev = nombre[::-1]
    nombre_truncado_rev = textwrap.shorten(nombre_rev, width=120, placeholder=" ]...[")
    global nombre_truncado
    nombre_truncado = textwrap.fill(nombre_truncado_rev[::-1], width=50)
    etiquetasaludo.configure(text="File path:\n" + nombre_truncado, font=(my_font2))
    selector.place(relx=0.37, rely=0.70)
    return rutaarchivo

def switch_boton(value):
    frame_abajo.configure(fg_color="#f0c808")
    frame_arriba.configure(fg_color="white")
    global botontransformar
    global progressbar
    global botonruta
    global opcion
    opcion = selector.get()
    etiquetasaludo.forget()
    global etiquetasaludo2
    etiquetasaludo2 = customtkinter.CTkLabel(frame_abajo, text="File path:\n" + nombre_truncado,
                                             font=(my_font2), corner_radius=55)
    etiquetasaludo2.pack(padx=50, pady=20)
    botontransformar = customtkinter.CTkButton(frame_abajo, text="Convert\n" + opcion, command=transformar, height=100, width=200, fg_color="#B31312", hover_color="#B34312", font=("helvetica", 20))
    botontransformar.pack(padx=50, pady=10)
    selector.place_forget()
    imagen1.forget()
    imagen2.pack(padx=20)
    frame_arriba.pack_propagate(True)
    frame_abajo.pack_propagate(True)

def transformar():
    global rutaarchivo
    global botonruta
    global progressbar
    botontransformar.configure(text="Converting\n"+ opcion)
    seleccion = selector.get()
    etiquetasaludo2.forget()
    progressbar = customtkinter.CTkProgressBar(frame_abajo, orientation="horizontal", mode="indeterminate", width=200, height=23, fg_color="white")
    progressbar.pack(padx=50, pady=20)
    progressbar.start()

    if seleccion == "PDF to Word":
        progressbar.start()
        label.pack(padx=50)
        rutapdf = rutaarchivo+".pdf"
        rutadocx = rutaarchivo+".docx"

        def pdf_to_docx(rutapdf, rutadocx):
            try:
                cv = Converter(rutapdf)
                cv.convert(rutadocx, start=0, end=None)
                cv.close()
                show_success_message()
            except Exception as e:
                progressbar.stop()
                progressbar.forget()  
                etiquetasaludo.configure(text="Error: " + str(e), fg_color="red") 
        threading.Thread(target=pdf_to_docx, args=(rutapdf, rutadocx)).start()

    elif seleccion == "PDF to Excel":

        progressbar.start()
        label.pack(padx=50, pady=30)
        rutapdf = rutaarchivo+".pdf"
        rutaxlsx = rutaarchivo+".xlsx"
        def pdf_to_excel(rutapdf, rutaxlsx):
            try:
                tables = camelot.read_pdf(rutapdf)

                for i, table in enumerate(tables):
                    # Create a unique Excel filename based on the table index
                    table.df.to_excel(rutaxlsx, index=False)  # Export to Excel
                show_success_message()

            except Exception as e:
                            progressbar.stop()
                            progressbar.forget()  
                            etiquetasaludo.configure(text="Error: " + str(e), fg_color="red") 
        threading.Thread(target=pdf_to_excel, args=(rutapdf, rutaxlsx)).start()

    elif seleccion == "PDF to CSV":
        rutapdf = rutaarchivo+".pdf"
        rutacsv = rutaarchivo+".csv"
        def pdf_to_csv(rutapdf, rutacsv):
                try:
                    tables = camelot.read_pdf(rutapdf)
                    tables.export(rutacsv, f='csv') # json, excel, html, markdown, sqlite
                    show_success_message() 

                except Exception as e:
                            progressbar.stop()
                            progressbar.forget()  
                            etiquetasaludo.configure(text="Error: " + str(e), fg_color="red") 
        threading.Thread(target=pdf_to_csv, args=(rutapdf, rutacsv)).start()  

    elif seleccion == "PDF to OCR":
        progressbar.start()
        label.configure(text="Applying OCR to your file...")
        label.pack(padx=50)
        rutapdf = rutaarchivo+".pdf"
        rutaocr = rutaarchivo+"_OCR"+".pdf"
        try:
            mi_hilo = threading.Thread(target=ocr_my_pdf, args=(rutapdf, rutaocr), daemon=True)
            mi_hilo.start()
            mi_hilo.join()

        except Exception as e:
            progressbar.stop()
            progressbar.forget()  
            etiquetasaludo.configure(text="Error: " + str(e), fg_color="red") 
    
    show_success_message()

    rutaarchivo = ""
    etiquetaruta.configure(text="")
    selector.set("Now Pick target format")

def show_checkmark():
    CTkMessagebox(title="Operation Finished", message="Conversion completed successfully", icon="check", option_1="Ok")
    label.forget()
    imagen1.pack(pady=20, padx=20)

def show_success_message():
    global botonruta 
    progressbar.stop()
    progressbar.forget()  
    show_checkmark()
    frame_arriba.configure(fg_color="#f0c808")
    frame_abajo.configure(fg_color="white")
    botontransformar.forget()
    etiquetasaludo.configure(text="Point your wand and choose a PDF \nfor alchemical transformation.", font=(my_font))
    etiquetasaludo.pack(padx=50, pady=20)
    botonruta.pack(padx=50)
    imagen1.pack(pady=20, padx=20)
    imagen2.forget()

def show_error_message(message):
    CTkMessagebox(title="Error", message=message, icon="cancel", option_1="Ok")

def ocr_my_pdf(rutapdf, rutasalida):
    error_log = {}
    try:
        # Perform OCR and create new PDF with extracted text
        output_file = ocrmypdf.ocr(
        rutapdf, rutasalida, output_type="pdf", skip_text=True, deskew=True
        )

        extraction_pdfs = {}
        pages_df = pd.DataFrame(columns=["text"])
        doc = fitz.open(output_file) # Open the newly created OCRed PDF
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            pages_df = pd.concat(
                [pages_df, pd.DataFrame([{"text": page.get_text("text")}])],
                ignore_index=True,
        )
        extraction_pdfs[output_file] = pages_df
        return extraction_pdfs
    except Exception as e:
        error_log[rutapdf] = str(e) # Convert exception to string for better logging

my_font = customtkinter.CTkFont(family=("helvetica"), size=24, weight="bold")
my_font2 = customtkinter.CTkFont(family=("helvetica"), size=16, weight="bold")

# Crear los frames
frame_arriba = customtkinter.CTkFrame(app, fg_color="#f0c808", height=50)
frame_abajo = customtkinter.CTkFrame(app, fg_color="white")

# Posicionar los frames (arriba - centro - abajo)
frame_arriba.pack(side=customtkinter.TOP, fill=tkinter.BOTH, expand=True)
frame_arriba.pack_propagate(False)
frame_abajo.pack(side=tkinter.BOTTOM, fill=tkinter.BOTH, expand=True)
frame_abajo.pack_propagate(False)
etiquetasaludo = customtkinter.CTkLabel(frame_arriba, text="First, choose the file you want \nto work your magic on...", font=(my_font), corner_radius=55)
etiquetasaludo.pack(padx=50, pady=30)

rutaimagen1 = os.path.abspath("image.jpeg")
rutaimagen2 = os.path.abspath("imagen2.jpg")

gato1 = customtkinter.CTkImage(light_image=Image.open(rutaimagen1),size=(500,269))
gato2 = customtkinter.CTkImage(light_image=Image.open(rutaimagen2),size=(400, 200))
# Creating label and packing it to the window.
imagen1 = customtkinter.CTkLabel(frame_abajo, image=gato1, text="")
imagen1.pack(pady=20, padx=20)
imagen2 = customtkinter.CTkLabel(frame_arriba, image=gato2, text="")

botonruta = customtkinter.CTkButton(frame_arriba, text="Select a PDF to start", command=seleccionar_documento, fg_color="#dd1c1a", hover_color="#ea4b48", width=50, height=50)
botonruta.pack(padx=50)

etiquetaruta = customtkinter.CTkLabel(frame_arriba, text="", font=("helvetica",15), text_color="#f0c808")
etiquetaruta.pack(padx=50)

# Opciones del selector
opciones = ["PDF to Excel", "PDF to Word", "PDF to CSV", "PDF to OCR"]

# Crear el selector
selector = customtkinter.CTkOptionMenu(frame_arriba, values=opciones, command=switch_boton)
selector.set("Now Pick target format")

txt = ("Brewing the elixir. Your PDF will be renewed!\nJust a little patience")
count = 0
text = ''

label = customtkinter.CTkLabel(frame_abajo,
                               text=txt,
                               font=('Helvetica', 20),
                               text_color="black"
                               )

app.mainloop()
