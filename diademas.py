# Instalar dependencias: "pip install pandas openpyxl docxtpl docx2pdf PyPDF2"

import getpass
import sys
import time
import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import os


PASSWORD = "0909"


pwd = getpass.getpass("CONTRASEÑA: ")

if pwd != PASSWORD:
    print("❌ Contraseña incorrecta. Acceso denegado.")
    sys.exit()


frames = [
r"""
 /\_/\  
( o.o )   
 > ^ <    
""",
r"""
  /\_/\ 
 ( o.o )  
  > ^ <   
""",
r"""
   /\_/\ 
  ( o.o ) 
   > ^ <  
""",
r"""
    /\_/\ 
   ( o.o )
    > ^ < 
"""
]


def clear():
    os.system('cls' if os.name == 'nt' else 'clear')


for i in range(2):  
    for frame in frames:
        clear()
        print(frame)
        time.sleep(0.2)


clear()
print(r"""
 /\_/\  
( o.o )  Hi!
 > ^ < 
""")

print(r"""
 __________.__                                .__    .___      
\______   \__| ____   _______  __ ____   ____ |__| __| _/____  
 |    |  _/  |/ __ \ /    \  \/ // __ \ /    \|  |/ __ |/  _ \ 
 |    |   \  \  ___/|   |  \   /\  ___/|   |  \  / /_/ (  <_> )
 |______  /__|\___  >___|  /\_/  \___  >___|  /__\____ |\____/ 
        \/        \/     \/          \/     \/        \/       

   ___         ___                 __  ___
  / _ )__ __  / _ \___ __  __(_)__/ / / _ )___   (_)__ ________ ____  ___ 
 / _  / // / / // / _ `/ |/ / / _  / / _  / -_) / / _ `/ __/ _ `/ _ \/ _ \
/____/\_, / /____/\_,_/|___/_/\_,_/ /____/\__/_/ /\_,_/_/  \_,_/_//_/\___/
     /___/                                  |___/                          
      """)

print("Cargando...\n")

df = pd.read_excel("actas.xlsx", header=None)


header_row = None
for i, row in df.iterrows():
    if "Nickname" in row.values:
        header_row = i
        break

if header_row is None:
    raise ValueError("No encontré la fila con 'Nickname'. Revisa el Excel.")


df = pd.read_excel("actas.xlsx", header=header_row)
print("Columna Corregida", df.columns.tolist())


plantilla = "Grupo_24_Agosto_21_2025-2.docx"
os.makedirs("pdfs", exist_ok=True)
pdfs = []

for i, row in df.iterrows():
    doc = DocxTemplate(plantilla)

  
    fecha = row["FechaEntrega"]
    if pd.notna(fecha):
        try:
            fecha = pd.to_datetime(fecha).strftime("%d/%m/%Y")  # Solo fecha
        except:
            fecha = str(fecha)

    contexto = {
        "nickname": row["Nickname"],
        "fecha": fecha,
        "nombre": row["Nombre"],
        "cc": row["CC"],
        "diadema": row["Diadema"]
    }

    doc.render(contexto)
    
    
    temp_docx = f"pdfs/temp_{i}.docx"
    doc.save(temp_docx)

    
    temp_pdf = temp_docx.replace(".docx", ".pdf")
    convert(temp_docx, temp_pdf)

   
    reader = PdfReader(temp_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        text = page.extract_text() or ""
        if text.strip():  
            writer.add_page(page)

    clean_pdf = f"pdfs/{row['Nickname']}.pdf"
    with open(clean_pdf, "wb") as f:
        writer.write(f)

    pdfs.append(clean_pdf)

   
    os.remove(temp_docx)
    os.remove(temp_pdf)

print("PDFs ✅")


merger = PdfMerger()
for pdf in pdfs:
    merger.append(pdf)

final_pdf = "Grupo_24_Agosto_21_2025-2.pdf"
merger.write(final_pdf)
merger.close()

print(f"Archivo final creado✅: {final_pdf}")
