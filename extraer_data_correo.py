import os, sys
import win32com.client
import extract_msg
import pdfplumber
import re
from openpyxl import Workbook, load_workbook

# 📂 Carpeta base = donde está el exe o script
if getattr(sys, 'frozen', False):  # ejecutado como exe
    BASE_DIR = os.path.dirname(sys.executable)
else:  # ejecutado como .py
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 📂 Carpeta para PDFs
PDF_DIR = os.path.join(BASE_DIR, "pdfs")
os.makedirs(PDF_DIR, exist_ok=True)

# 📊 Archivo Excel donde guardaremos los resultados
excel_file = os.path.join(BASE_DIR, "ordenes_compra.xlsx")

# Si no existe, crear Excel con encabezados
if not os.path.exists(excel_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ordenes"
    ws.append(["Buzón", "Número de Orden", "Sección", "Total sin IVA", "Total con IVA", "Descripción final"])
    wb.save(excel_file)

# 📧 Conectarse a Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as e:
    print("❌ Error al conectarse a Outlook:", e)
    sys.exit()

# Filtro por asunto
filtro = "@SQL=\"urn:schemas:mailheader:subject\" like '%Cencosud: Orden de Compra%'"

# 🔄 Recorremos todos los buzones
for store in outlook.Folders:
    print(f"📂 Revisando buzón: {store.Name}")

    try:
        inbox = store.Folders["Bandeja de entrada"]
        correos_filtrados = inbox.Items.Restrict(filtro)
    except Exception as e:
        print(f"⚠️ No se pudo acceder a la bandeja de entrada de {store.Name}: {e}")
        continue

    for mail in correos_filtrados:
        print("   📌 Asunto:", mail.Subject)

        try:
            for attachment in mail.Attachments:
                if attachment.FileName.lower().endswith(".msg"):
                    msg_path = os.path.join(BASE_DIR, attachment.FileName)
                    attachment.SaveAsFile(msg_path)

                    msg = extract_msg.Message(msg_path)

                    for att in msg.attachments:
                        if att.longFilename.lower().endswith(".pdf"):
                            pdf_path = os.path.join(PDF_DIR, att.longFilename)
                            with open(pdf_path, "wb") as f:
                                f.write(att.data)

                            # ---- Leer y procesar PDF ----
                            with pdfplumber.open(pdf_path) as pdf:
                                texto = ""
                                for page in pdf.pages:
                                    texto += page.extract_text() + "\n"

                            # 1️⃣ Número orden
                            match_orden = re.search(r"Número\s*(?:de\s*)?orden\s*[:\s]+(\d+)", texto, re.IGNORECASE)
                            numero_orden = match_orden.group(1) if match_orden else "NO ENCONTRADO"

                            # 2️⃣ Sección
                            match_seccion = re.search(r"Sección\s*:\s*(.+)", texto, re.IGNORECASE)
                            seccion = match_seccion.group(1).strip() if match_seccion else "NO ENCONTRADO"

                            # 3️⃣ Total sin IVA
                            neto = re.search(r"Neto total sin IVA\s+[A-Z]{3}\s+([\d.,]+)", texto)
                            total_sin_iva = neto.group(1) if neto else None

                            # 4️⃣ Total con IVA (última coincidencia)
                            totales = re.findall(r"Total\s+([\d.,]+)", texto)
                            total_con_iva = totales[-1] if totales else None

                            # 5️⃣ Línea después de "Total"
                            lineas = [l.strip() for l in texto.splitlines() if l.strip()]
                            descripcion_final = None
                            for i, linea in enumerate(lineas):
                                if linea.startswith("Total") and i + 1 < len(lineas):
                                    descripcion_final = lineas[i + 1]
                                    break

                            # 📊 Guardar en Excel
                            wb = load_workbook(excel_file)
                            ws = wb.active
                            ws.append([store.Name, numero_orden, seccion, total_sin_iva, total_con_iva, descripcion_final])
                            wb.save(excel_file)

                    msg.close()
                    os.remove(msg_path)  # limpiar el .msg temporal
        except Exception as e:
            print(f"⚠️ Error procesando correo en {store.Name}: {e}")
