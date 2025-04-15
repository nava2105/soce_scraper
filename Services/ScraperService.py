import requests
import json
import pandas as pd
import re
import time
from bs4 import BeautifulSoup


def fetch_admin(modified_links):
    data = []
    for i, link in enumerate(modified_links, start=1):
        link = link.replace("ProcesoContratacion/tab.php?tab=1&id=", "EC/resumenContractual1.cpe?idSoliCompra=")
        response = requests.get(link)

        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            tds = soup.find_all('td')

            if len(tds) > 86:
                codigo_proceso = tds[7].text.strip()
                administrador_contrato = tds[85].text.strip()
                data.append({
                    "Código de Proceso": codigo_proceso,
                    "Administrador de Contrato": administrador_contrato,
                    "Link": link
                })
                print(i, " | ",data)
        else:
            print(f"Failed to retrieve the page {link}. Status code: {response.status_code}")
    return data

def fetch_provider(modified_links):
    data = []
    for i, link in enumerate(modified_links, start=1):
        link = link.replace("ProcesoContratacion/tab.php?tab=1&id=", "EC/resumenContractual1.cpe?idSoliCompra=")
        response = requests.get(link)

        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            tds = soup.find_all('td')
            soup2 = BeautifulSoup(tds[38].decode_contents(), 'html.parser')
            tds2 = soup2.find_all('td')

            if len(tds) > 86:
                codigo_proceso = tds[7].text.strip()
                tipo_compra = tds[16].text.strip()
                presupuesto_referencial = tds[19].text.strip()
                if len(tds2) > 6:
                    ruc_proveedor = tds2[6].text.strip()
                    nombre_adjudicatario = tds2[7].text.strip()
                    fecha_adjudicacion = tds2[8].text.strip()
                    monto_adjudicacion = tds2[9].text.strip()
                    data.append({
                        "Código de Proceso": codigo_proceso,
                        "Tipo de Compra": tipo_compra,
                        "Presupuesto Referencial": presupuesto_referencial,
                        "Ruc Proveedor": ruc_proveedor,
                        "Nombre Adjudicatario": nombre_adjudicatario,
                        "Fecha Adjudicacion": fecha_adjudicacion,
                        "Monto Adjudicacion": monto_adjudicacion,
                        "Link": link
                    })
                else:
                    data.append({
                        "Código de Proceso": codigo_proceso,
                        "Tipo de Compra": tipo_compra,
                        "Presupuesto Referencial": presupuesto_referencial,
                        "Ruc Proveedor": "",
                        "Nombre Adjudicatario": "",
                        "Fecha Adjudicacion": "",
                        "Monto Adjudicacion": "",
                        "Link": link
                    })
                print(i, " | ", data)
        else:
            print(f"Failed to retrieve the page {link}. Status code: {response.status_code}")
    return data

class ScraperService:
    def __init__(self):
        self.url = "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/servicio/interfazWeb.php"

    def scrape(self, headers, data):

        all_data = []  # List to store all data

        # Iterate over the pages
        pagina_actual = 0
        while True:
            data["paginaActual"] = str(pagina_actual)
            response = requests.post(self.url, headers=headers, data=data)

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
