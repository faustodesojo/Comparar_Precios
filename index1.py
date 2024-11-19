import aiohttp
import asyncio
from bs4 import BeautifulSoup
import openpyxl

# Cargar el archivo de Excel
archivo_excel = r"C:\Users\Windows\Desktop\PYTHON\COMPARAR_PRECIOS.xlsm"
wb = openpyxl.load_workbook(archivo_excel)
ws = wb.active

# Función para procesar una sola URL
async def procesar_url(session, url, fila):
    try:
        async with session.get(url) as response:
            html = await response.text()
            soup = BeautifulSoup(html, "html.parser")
            # Extraer precio
            span_precio = soup.find("span", class_="price")
            if span_precio:
                precio_texto = (
                    span_precio.text.strip()
                    .replace("$", "")
                    .replace(".", "")
                    .replace(",", ".")
                )
                precio_numerico = float(precio_texto)
                ws.cell(row=fila, column=12, value=precio_numerico)  # Columna L
            else:
                ws.cell(row=fila, column=12, value="Precio no encontrado")
    except Exception as e:
        ws.cell(row=fila, column=12, value="Error")
        print(f"Error al procesar {url}: {e}")

# Función principal para manejar múltiples solicitudes
async def main():
    urls = [(fila[0].value, fila[0].row) for fila in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=13, max_col=13) if fila[0].value]
    async with aiohttp.ClientSession() as session:
        tareas = [procesar_url(session, url, fila) for url, fila in urls]
        await asyncio.gather(*tareas)

# Ejecutar
asyncio.run(main())

# Guardar el archivo Excel
wb.save(r"C:\Users\Windows\Desktop\PYTHON\COMPARAR_PRECIOS.xlsx")
print("Actualización completada. Archivo guardado.")
