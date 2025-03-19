import requests
import json
import pandas as pd
import re
import time

# URL we want to scrape from
url = "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/servicio/interfazWeb.php"

# Headers to be sent with the request
headers = {
    "Accept": "text/javascript, text/html, application/xml, text/xml, */*",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://www.compraspublicas.gob.ec",
    "Referer": "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/IC/buscarInfima.cpe",
}

# Function to make the request and process the data
def fetch_data(start_date, end_date):
    # Data to be sent in the POST request
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
        "f_inicio": start_date,
        "f_fin": end_date,
        "count": "1633",
        "paginaActual": "0",
        "estado": "",
        "trx": "",
        "_": ""
    }

    all_data = []  # List to store all data

    # Iterate over the pages
    pagina_actual = 0
    while True:
        data["paginaActual"] = str(pagina_actual)
        response = requests.post(url, headers=headers, data=data)

        # Verify that the request was successful
        if response.status_code == 200:
            headers_json = dict(response.headers)

            if 'X-JSON' in headers_json:
                try:
                    normalized_data = re.sub(r'[()]', '', headers_json['X-JSON'])
                    json_data = json.loads(normalized_data)
                    print(pagina_actual, ": ", json_data)

                    if not json_data:  # If there is no more data, exit the loop.
                        break

                    all_data.extend(json_data)  # Add data to the list
                    pagina_actual += 20  # Go to next page

                except json.JSONDecodeError:
                    print("Error: The content of 'X-JSON' is not a valid JSON.")
                    break
            else:
                print("The 'X-JSON' header was not found.")
                break
        else:
            print(f"Error in the application: {response.status_code}")
            break

        # Wait some time before the next request to avoid blockages.
        time.sleep(1)  # Wait 1 second

    return all_data

# Define date range
start_date = "2023-01-01"
end_date = "2025-03-16"

print(f"Fetching data from {start_date} to {end_date}...")
records = fetch_data(start_date, end_date)

# Convert to a Pandas DataFrame
df = pd.DataFrame(records)

# Rename columns
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

# Delete the "no" and "i" columns if they exist
df.drop(columns=["no", "i"], errors='ignore', inplace=True)

# Convert "Quantity" and "U. Cost" to numeric type
df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors='coerce')
df["Costo U."] = pd.to_numeric(df["Costo U."], errors='coerce')

# Create the "Value" column by multiplying "Quantity" by "U. Cost".
df["Valor"] = df["Cantidad"] * df["Costo U."]

# Show the first records
print(df.head())

# Save to Excel file
df.to_excel(f"INF-{start_date}-{end_date}.xlsx", index=False)

print("Data successfully exported to Excel.")