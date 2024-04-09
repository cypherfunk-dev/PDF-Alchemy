import customtkinter
from CTkMessagebox import CTkMessagebox
import tkinter
import os
import threading
import textwrap
from pdf2docx import Converter
import tabula
import pandas as pd
from tabula.io import read_pdf
from PIL import Image, ImageTk

jvm_options = ["-XX:ReservedCodeCacheSize=128m"]

# Estilo e inicialización
customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("blue")

app = customtkinter.CTk()
app.geometry("600x500")
app.title('PDF Alchemy - Convert PDF to Word/Excel')
app.resizable(False, False)










# Definición de variables globales
rutaarchivo = ""
botonruta = None
global imagenql2

# Definición de funciones
def seleccionar_documento():
    global rutaarchivo
    global botonruta
    filename = customtkinter.filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
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
    opcion = selector.get()
    etiquetasaludo.forget()
    global etiquetasaludo2
    etiquetasaludo2 = customtkinter.CTkLabel(frame_abajo, text="File path:\n" + nombre_truncado,
                                             font=(my_font2), corner_radius=55)
    etiquetasaludo2.pack(padx=50, pady=20)
    botontransformar = customtkinter.CTkButton(frame_abajo, text="Convert\n" + opcion, command=transformar, height=100, width=200, fg_color="#B31312", hover_color="#B34312", font=("helvetica", 20))
    botontransformar.pack(padx=50, pady=10)
    selector.place_forget()
    imagenql.forget()
    imagenql2.pack(padx=20)


def transformar():
    global rutaarchivo
    global botonruta
    seleccion = selector.get()
    etiquetasaludo2.forget()
    progressbar = customtkinter.CTkProgressBar(frame_abajo, orientation="horizontal", mode="indeterminate", width=200, height=23, fg_color="white")
    progressbar.pack(padx=50, pady=20)
    
    def show_success_message(message):
        global botonruta 
        progressbar.stop()
        progressbar.forget()  
        show_checkmark()
        frame_arriba.configure(fg_color="#f0c808")
        frame_abajo.configure(fg_color="white")
        botontransformar.forget()
        etiquetasaludo.configure(text="Point your wand! \nChoose a PDF for alchemical transformation.")
        etiquetasaludo.pack(padx=50, pady=20)
        botonruta = customtkinter.CTkButton(frame_arriba, text="Select a PDF to start", command=seleccionar_documento, fg_color="#dd1c1a", hover_color="#ea4b48", width=50, height=50)
        botonruta.pack(padx=50)
        imagenql.pack(pady=20, padx=20)
        imagenql2.forget()


    if seleccion == "PDF to Word":
        progressbar.start()
        label.pack(padx=50, pady=30)
        rutapdf = rutaarchivo+".pdf"
        rutadocx = rutaarchivo+".docx"

        def pdf_to_docx(rutapdf, rutadocx):
            try:
                cv = Converter(rutapdf)
                cv.convert(rutadocx, start=0, end=None)
                cv.close()
                show_success_message("Conversion to Word completed successfully") 
            except Exception as e:
                progressbar.stop()
                progressbar.forget()  
                # etiquetasaludo.configure(text="Error: " + str(e), fg_color="red") 
        threading.Thread(target=pdf_to_docx, args=(rutapdf, rutadocx)).start()

    elif seleccion == "PDF to Excel":
        progressbar.start()
        label.pack(padx=50, pady=30)
        rutapdf = rutaarchivo+".pdf"
        rutaxlsx = rutaarchivo+".xlsx"

        def pdf_to_excel(rutapdf, rutaxlsx):
            try:
                tables = tabula.read_pdf(rutapdf, pages='all', java_options=jvm_options)
                with pd.ExcelWriter(rutaxlsx) as writer:
                    for i, table in enumerate(tables):
                        table.to_excel(writer, sheet_name=f'Sheet{i+1}')
                show_success_message("Conversion to Excel completed successfully") 
            except Exception as e:
                progressbar.stop()
                progressbar.forget()  
                # etiquetasaludo.configure(text="Error: " + str(e), fg_color="red") 

        threading.Thread(target=pdf_to_excel, args=(rutapdf, rutaxlsx)).start()

    else:
        print("Error")

    rutaarchivo = ""
    etiquetaruta.configure(text="")
    selector.set("Now Pick target format")

def show_checkmark():
    CTkMessagebox(title="Operation Finished", message="Conversion completed successfully", icon="check", option_1="Ok")
    label.forget()
    imagenql.pack(pady=20, padx=20)



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

gato1 = customtkinter.CTkImage(light_image=Image.open("image.jpeg"),size=(500,269))
gato2 = customtkinter.CTkImage(light_image=Image.open("imagen2.jpg"),size=(1021,508))
# Creating label and packing it to the window.
imagenql = customtkinter.CTkLabel(frame_abajo, image=gato1, text="")
imagenql.pack(pady=20, padx=20)
imagenql2 = customtkinter.CTkLabel(frame_arriba, image=gato2, text="")

botonruta = customtkinter.CTkButton(frame_arriba, text="Select a PDF to start", command=seleccionar_documento, fg_color="#dd1c1a", hover_color="#ea4b48", width=50, height=50)
botonruta.pack(padx=50)

etiquetaruta = customtkinter.CTkLabel(frame_arriba, text="", font=("helvetica",15), text_color="#f0c808")
etiquetaruta.pack(padx=50)

# Opciones del selector
opciones = ["PDF to Excel", "PDF to Word"]

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
