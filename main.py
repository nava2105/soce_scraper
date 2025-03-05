import requests
import json
import pandas as pd
import re

# URL de la que queremos hacer scraping
url = "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/servicio/interfazWeb.php"

# Encabezados que se deben enviar con la solicitud
headers = {
    "Accept": "text/javascript, text/html, application/xml, text/xml, */*",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://www.compraspublicas.gob.ec",
    "Referer": "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/IC/buscarInfima.cpe",
}

# Datos que se enviarán en la solicitud POST
data = {
    "__class": "TcomInfima",
    "__action": "buscarInfimaxEntidad",
    "txtEntidadContratante": "CORPORACION ELECTRICA DEL ECUADOR CELEC EP.",
    "cmbEntidad": "238940",
    "txtNroFactInfima": "",
    "txtIdProducto": "",
    "txtObsInfima": "",
    "txtCodTipoCompraTab": "176",
    "txtCodTipoCompra": "",
    "txtMes": "0",
    "txtAnio": "0",
    "f_inicio": "2024-09-5",
    "f_fin": "2025-03-5",
    "count": "1633",
    "paginaActual": "0",
    "estado": "",
    "trx": "",
    "_": ""
}

# Realizar la solicitud POST
response = requests.post(url, headers=headers, data=data)

# Verificar que la solicitud fue exitosa
if response.status_code == 200:
    # Obtener los encabezados de la respuesta
    headers_json = dict(response.headers)

    # Verificar si 'X-JSON' está en los encabezados
    if 'X-JSON' in headers_json:
        try:
            normalized_data = re.sub(r'[()]', '', headers_json['X-JSON'])
            print(normalized_data)
            # Convertir el contenido de 'X-JSON' a un diccionario de Python
            json_data = json.loads(normalized_data)

            # Convertir a un DataFrame de Pandas
            df = pd.DataFrame(json_data)

            df.rename(columns={
                "nu": "Nro.",
                "n": "Nro.Factura",
                "f": "Fecha de emisión de la factura",
                "c": "CPC",
                "d": "Descripción CPC",
                "r": "Razón Social",
                "co": "Objeto de Compra",
                "ca": "Cantidad",
                "t": "Costo U.",
                "j": "Justificativo",
                "ct": "Tipo de Compra",
                "nom": "Responsable de Asuntos Administrativos"
            }, inplace=True)

            # Eliminar las columnas "no" e "i" si existen
            df.drop(columns=["no", "i"], errors='ignore', inplace=True)

            # Convertir "Cantidad" y "Costo U." a tipo numérico
            df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors='coerce')
            df["Costo U."] = pd.to_numeric(df["Costo U."], errors='coerce')

            # Crear la columna "Valor" multiplicando "Cantidad" por "Costo U."
            df["Valor"] = df["Cantidad"] * df["Costo U."]

            # Mostrar los primeros registros
            print(df.head())

            # Guardar en un archivo Excel
            df.to_excel("datos_extraidos.xlsx", index=False)

            print("Datos exportados correctamente a 'datos_extraidos.xlsx'.")

        except json.JSONDecodeError:
            print("Error: El contenido de 'X-JSON' no es un JSON válido.")
    else:
        print("No se encontró el encabezado 'X-JSON'.")
else:
    print(f"Error en la solicitud: {response.status_code}")
