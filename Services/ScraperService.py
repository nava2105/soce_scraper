import requests
import json
import pandas as pd
import re
import time

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