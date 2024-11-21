import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from pandas.api.types import is_datetime64_any_dtype

# Configuración de Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales_json = r"C:\Users\Windows\Desktop\PYTHON\lista-505\absolute-cache-442317-a3-e81f25d2fc0f.json"

credentials = ServiceAccountCredentials.from_json_keyfile_name(credenciales_json, scope)
client = gspread.authorize(credentials)

# Abrir el archivo de Google Sheets
spreadsheet_id = "1JoxqdX0wnVLJ5kD0T9AXOGf13g8RbrKAihajpUq53sw"
spreadsheet = client.open_by_key(spreadsheet_id)

# Seleccionar la hoja número 7
sheet = spreadsheet.get_worksheet(6)  # Índice 6 = hoja 7

# Leer el archivo Excel
excel_path = r"Z:\base\lista0505.xls"
excel_data = pd.read_excel(excel_path, sheet_name="DEPOSITO")

# Convertir columnas de fecha/hora a texto
for column in excel_data.columns:
    if is_datetime64_any_dtype(excel_data[column]):
        excel_data[column] = excel_data[column].astype(str)

# Reemplazar valores NaN con una cadena vacía
excel_data = excel_data.fillna('')

# Convertir los datos a una lista y escribirlos en la hoja de Google Sheets
rows = excel_data.values.tolist()
sheet.clear()  # Limpia la hoja antes de escribir nuevos datos
sheet.append_rows(rows)

print("Hoja 'DEPOSITO' sincronizada con Google Sheets.")
