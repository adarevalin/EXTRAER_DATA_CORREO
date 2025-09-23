import os, sys
import win32com.client
import extract_msg
import pdfplumber
import re
from openpyxl import Workbook, load_workbook
from datetime import datetime

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
    ws.append([
        "Buzón", 
        "Número de Orden", 
        "Fecha Creación", 
        "Fecha Entrega", 
        "Sección", 
        "Total sin IVA", 
        "Total con IVA", 
        "Descripción final"
    ])
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
        correos_filtrados.Sort("[ReceivedTime]", True)  # ordenar descendente
        correos_a_procesar = list(correos_filtrados)[:700]  # tomar los 700 más recientes
    except Exception as e:
        print(f"⚠️ No se pudo acceder a la bandeja de entrada de {store.Name}: {e}")
        continue

    for mail in correos_a_procesar:
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

                            # 2️⃣ Fecha creación
                            match_creacion = re.search(r"Fecha\s*creaci[oó]n\s*[:\s]+([\d./-]+)", texto, re.IGNORECASE)
                            if match_creacion:
                                try:
                                    fecha_creacion = datetime.strptime(match_creacion.group(1), "%d.%m.%Y").strftime("%d/%m/%Y")
                                except:
                                    fecha_creacion = match_creacion.group(1)
                            else:
                                fecha_creacion = "NO ENCONTRADO"

                            # 3️⃣ Fecha entrega
                            match_entrega = re.search(r"Fecha\s*entrega\s*[:\s]+([\d./-]+)", texto, re.IGNORECASE)
                            if match_entrega:
                                try:
                                    fecha_entrega = datetime.strptime(match_entrega.group(1), "%d.%m.%Y").strftime("%d/%m/%Y")
                                except:
                                    fecha_entrega = match_entrega.group(1)
                            else:
                                fecha_entrega = "NO ENCONTRADO"

                            # 4️⃣ Sección
                            match_seccion = re.search(r"Sección\s*:\s*(.+)", texto, re.IGNORECASE)
                            seccion = match_seccion.group(1).strip() if match_seccion else "NO ENCONTRADO"

                            # 5️⃣ Total sin IVA
                            neto = re.search(r"Neto total sin IVA\s+[A-Z]{3}\s+([\d.,]+)", texto)
                            total_sin_iva = neto.group(1) if neto else None

                            # 6️⃣ Total con IVA (última coincidencia)
                            totales = re.findall(r"Total\s+([\d.,]+)", texto)
                            total_con_iva = totales[-1] if totales else None

                            # 7️⃣ Línea después de "Total"
                            lineas = [l.strip() for l in texto.splitlines() if l.strip()]
                            descripcion_final = None
                            for i, linea in enumerate(lineas):
                                if linea.startswith("Total") and i + 1 < len(lineas):
                                    descripcion_final = lineas[i + 1]
                                    break

                            # 📊 Guardar en Excel solo si NO existe ya
                            wb = load_workbook(excel_file)
                            ws = wb.active

                            ordenes_existentes = [str(c.value) for c in ws["B"] if c.value]  # Columna B = Número de Orden

                            if numero_orden not in ordenes_existentes:
                                ws.append([
                                    store.Name, 
                                    numero_orden, 
                                    fecha_creacion, 
                                    fecha_entrega, 
                                    seccion, 
                                    total_sin_iva, 
                                    total_con_iva, 
                                    descripcion_final
                                ])
                                wb.save(excel_file)
                                #print(f"   ✅ Guardada orden {numero_orden}")
                            else:
                                print(f"   ⏭️ Orden {numero_orden} ya existe, saltada.")

                    msg.close()
                    os.remove(msg_path)  # limpiar el .msg temporal
        except Exception as e:
            print(f"⚠️ Error procesando correo en {store.Name}: {e}")


