from flask import Flask, request, Response, render_template, send_file
from datetime import datetime
import pandas as pd
import os
import re

from Services.ExcelService import export_excel_process_inf, export_excel_process_pro, export_excel_process_reg, export_excel_process_pre, load_hyperlinks_from_excel, save_dataframe_to_excel
from Services.LlmService import configure_gemini_api, generate_text_embeddings, generate_ai_response
from Services.PdfService import allowed_file, extract_text_chunks
from Services.ScraperService import ScraperService, fetch_admin
from Services.StorageService import delete_files_in_directory


scraper = ScraperService()
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
RESOLUTIONS_FOLDER = "resolutions"
JSON_STORE = 'embeddings_store.json'
RESPONSES_STORE = 'responses_store.json'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(RESOLUTIONS_FOLDER, exist_ok=True)

@app.route('/')
def index():
    delete_files_in_directory(OUTPUT_FOLDER)
    delete_files_in_directory(UPLOAD_FOLDER)
    return render_template('index.html')

@app.route('/processes')
def processes():
    delete_files_in_directory(OUTPUT_FOLDER)
    delete_files_in_directory(UPLOAD_FOLDER)
    return render_template('processes.html')

@app.route('/admins')
def admins():
    delete_files_in_directory(OUTPUT_FOLDER)
    delete_files_in_directory(UPLOAD_FOLDER)
    return render_template('admins.html')

@app.route('/technical_commissions')
def technical_commissions():
    delete_files_in_directory(OUTPUT_FOLDER)
    delete_files_in_directory(UPLOAD_FOLDER)
    return render_template('technical_commissions.html')

@app.route('/extract', methods=['POST'])
def extract():
    data = request.form
    print(data)
    start_date = str(data.get('start_date'))
    end_date = str(data.get('end_date'))
    type = int(data.get('type'))

    if type == 1: # Ínfima Cuantía
        headers = {
            "Accept": "text/javascript, text/html, application/xml, text/xml, */*",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://www.compraspublicas.gob.ec",
            "Referer": "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/IC/buscarInfima.cpe",
        }
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
            "paginaActual": "0",
            "estado": "",
            "trx": "",
            "_": ""
        }
        records = scraper.scrape(headers, data)
        filename = export_excel_process_inf(records, start_date, end_date)
    elif type == 2: # Procesos
        headers = {
            "Accept": "text/javascript, text/html, application/xml, text/xml, */*",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://www.compraspublicas.gob.ec",
            "Referer": "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/PC/buscarProceso.cpe?sg=1",
        }
        data = {
            "__class": "SolicitudCompra",
            "__action": "buscarProcesoxEntidad",
            "txtEntidadContratante": "CORPORACION ELECTRICA DEL ECUADOR CELEC EP.",
            "cmbEntidad": "238940",
            "txtCodigoTipoCompra": "",
            "txtCodigoProceso": "",
            "txtNroFactInfima": "",
            "f_inicio": start_date,
            "f_fin": end_date,
            "paginaActual": "0",
            "estado": "",
            "trx": "",
            "_": ""
        }
        records = scraper.scrape(headers, data)
        filename = export_excel_process_pro(records, start_date, end_date)
    elif type == 3: # Regimen Especial
        headers = {
            "Accept": "text/javascript, text/html, application/xml, text/xml, */*",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://www.compraspublicas.gob.ec",
            "Referer": "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/PC/buscarProcesoRE.cpe?op=R",
        }
        data = {
            "__class": "SolicitudCompra",
            "__action": "buscarProcesoxEntidad",
            "txtEntidadContratante": "CORPORACION ELECTRICA DEL ECUADOR CELEC EP.",
            "cmbEntidad": "238940",
            "txtTiposContratacion": "186",
            "txtCodTipoCompra": "",
            "txtCodigoProceso": "",
            "f_inicio": start_date,
            "f_fin": end_date,
            "paginaActual": "0",
            "estado": "",
            "trx": "50008",
            "_": ""
        }
        records = scraper.scrape(headers, data)
        filename = export_excel_process_reg(records, start_date, end_date)
    elif type == 4: # Procedimiento especial
        headers = {
            "Accept": "text/javascript, text/html, application/xml, text/xml, */*",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://www.compraspublicas.gob.ec",
            "Referer": "https://www.compraspublicas.gob.ec/ProcesoContratacion/compras/PC/buscarProcesoRE.cpe?op=P",
        }
        data = {
            "__class": "SolicitudCompra",
            "__action": "buscarProcesoxEntidad",
            "txtEntidadContratante": "CORPORACION ELECTRICA DEL ECUADOR CELEC EP.",
            "cmbEntidad": "238940",
            "txtTiposContratacion": "219",
            "txtCodTipoCompra": "",
            "txtCodigoProceso": "",
            "f_inicio": start_date,
            "f_fin": end_date,
            "paginaActual": "0",
            "estado": "",
            "trx": "50008",
            "_": ""
        }
        records = scraper.scrape(headers, data)
        filename = export_excel_process_pre(records, start_date, end_date)
    else:
        return Response("Incorrect type, try again", status=400)
    return send_file(filename, as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return Response("No file part", status=400)

    file = request.files['file']

    if file.filename == '':
        return Response("No selected file", status=400)

    if not file.filename.endswith('.xlsx'):
        return Response("Only Excel files are allowed", status=400)

    # Save the uploaded file temporarily
    file_path = os.path.join("uploads", file.filename)
    file.save(file_path)

    # Load hyperlinks from the uploaded Excel file
    hyperlinks = load_hyperlinks_from_excel(file_path)

    # Modify links as needed
    modified_links = [re.sub("PC/informacion", "", link).replace("2.cpe?idSoliCompra=", "/tab.php?tab=1&id=") for link
                      in hyperlinks]

    # Fetch details
    data = fetch_admin(modified_links)

    # Create DataFrame and save to Excel
    df = pd.DataFrame(data)
    output_file_path = save_dataframe_to_excel(df, file.filename)

    return send_file(output_file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
