import camelot # encardo de leer el pdf
import os # mapeo de archivo y directorio
import pandas as pd # libreria encargada de guardar el excel
from tkinter import Tk # crea una interfaz para seleccionar los activos y carpetas de una manera visual
from tkinter.filedialog import askopenfilename, askdirectory
from datetime import datetime

def convertir_pdf_a_excel():
    # se oculta la ventana principal de tkinter
    Tk().withdraw()

    # seleccionar el archivo pdf (abre una ventana para seleccionar el PDF)
    print("seleccione el archivo PDF:")
    pdf_file = askopenfilename(
        title = "seleccionar el archivo PDF",
        filetypes = [("archivo PDF", "*.pdf")]
    )
    if not pdf_file:
        print("No se selecciono ningun archivo. Saliendo...")
        return
    
    #Seleccionar carpeta de destino
    print("seleccione la carpeta de destino del archivo:")
    output_folder = askdirectory(title = "Seleccionar carpeta")
    if not output_folder:
        print("no se selecciono la carpeta de destino. Saliendo...")
        return
    
    #  exel importado nombre unico del archivo con fecha y hora con la extension XLSX
    now = datetime.now().strftime("%m-%d-%y %H %M")
    output_excel = os.path.join(output_folder, f"Excel Importado {now}.xlsx")

    # extrar tablas PDF
        # camelot busca alguna tabla si no hay se cierra
    print("extrayendo tablas del PDF...")

    tables = camelot.read_pdf(pdf_file, pages = "all", flavor = "stream")
    if not tables:
        print("No se encontraron tablas en el archivo PDF:")
        return
    
    # guardar todas las tablas en un solo archivo excel
        # si hay una tabla se usa panda
    with pd.ExcelWriter(output_excel, engine = "openpyxl") as writer:
        for i, table in enumerate(tables):
            table.df.to_excel(writer, sheet_name = f"Tabla_{i+1}", index = False)

    #se ejecuta la funcion
if __name__ == "__main__":
        convertir_pdf_a_excel()