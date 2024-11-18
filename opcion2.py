import xlwings as xw
import requests
from bs4 import BeautifulSoup

# Abrir el archivo de Excel con xlwings
archivo_excel = r"C:\Users\Windows\Desktop\PYTHON\COMPARAR_PRECIOS.xlsm"
app = xw.App(visible=False)  # No mostrar la ventana de Excel
wb = app.books.open(archivo_excel)
ws = wb.sheets[0]  # Asumiendo que los datos están en la primera hoja

# Iterar sobre las filas para obtener URLs
for fila in range(2, ws.cells.last_cell.row + 1):  # Comienza en la fila 2 para omitir encabezados
    url = ws.cells(fila, 13).value  # Columna M (13) para las URLs
    if url:
        try:
            # Hacer la solicitud a la URL
            response = requests.get(url)
            response.raise_for_status()  # Asegurarse de que la solicitud no devuelva errores HTTP
            soup = BeautifulSoup(response.content, "html.parser")

            # Extraer el precio del span con la clase "price"
            span_precio = soup.find("span", class_="price")  # Ajusta la clase según la página
            precio = span_precio.text.strip() if span_precio else "No encontrado"

            # Guardar el precio en la columna L (12)
            ws.cells(fila, 12).value = precio

        except Exception as e:
            print(f"Error al procesar {url}: {e}")
            ws.cells(fila, 12).value = "Error"  # Escribe 'Error' en caso de fallo

# Guardar los cambios en el archivo Excel
wb.save()
wb.close()
app.quit()  # Cerrar la aplicación de Excel

print("Actualización completada. Archivo guardado.")
