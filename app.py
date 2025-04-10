from flask import Flask, request, Response, render_template, send_file, redirect, url_for, jsonify
from werkzeug.utils import secure_filename
from datetime import datetime
import pandas as pd
import os
import re
import csv
import io
import json
import datetime

from Services.ExcelService import export_excel_process_inf, export_excel_process_pro, export_excel_process_reg, export_excel_process_pre, load_hyperlinks_from_excel, save_dataframe_to_excel
from Services.LlmService import configure_gemini_api, generate_text_embeddings, generate_ai_response
from Services.PdfService import allowed_file, extract_text_chunks
from Services.ScraperService import ScraperService, fetch_admin
from Services.StorageService import delete_files_in_directory


scraper = ScraperService()
app = Flask(__name__)
configure_gemini_api()

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
RESOLUTIONS_FOLDER = "resolutions"
JSON_STORE = 'embeddings_store.json'
RESPONSES_STORE = 'responses_store.json'
prompt_template = """En base a las resoluciones o el apartado donde se resuelve a los miembros del comité técnico, lista a cada miembro y su función tanto en la empresa como en el comité.
-Contexto: Se necesita extraer información de los miembros del comité técnico del proceso para alimentar las bases de datos.
-Instrucción: Extrae la información de todos los miembros del comité técnico y sus cargos tanto para con el comité como para con la empresa
-Formato: Muestra la información solicitada en formato JSON, siguiendo el orden de Miembro del comité, Cargo en la empresa, Cargo en la comisión
-Restricciones: Si no encuentras la información en el texto no la pongas o no te inventes nombres, y no pongas el nombre de la persona que firmó el documento, además si no encuentras el cargo dentro de la comisión, repite en ese campo el cargo en la empresa pero no puede decir simplemente miembro, y si no encuentras el cargo en la empresa repite el cargo en el comité"""

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(RESOLUTIONS_FOLDER, exist_ok=True)

def load_json_store(filepath):
    if os.path.exists(filepath):
        with open(filepath, 'r') as file:
            return json.load(file)
    return {}


def save_json_store(filepath, data):
    with open(filepath, 'w') as file:
        json.dump(data, file, indent=4)


embeddings_store = load_json_store(JSON_STORE)
responses_store = load_json_store(RESPONSES_STORE)
configure_gemini_api()
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
    page = request.args.get('page', 1, type=int)
    search_query = request.args.get('search', '', type=str)
    per_page = 10
    files = list(embeddings_store.keys())
    if search_query:
        files = [file for file in files if search_query.lower() in file.lower()]
    total_files = len(files)
    total_pages = (total_files + per_page - 1) // per_page

    start = (page - 1) * per_page
    end = start + per_page
    paginated_files = files[start:end]
    return render_template("technical_commissions.html", files=paginated_files, total_pages=total_pages, current_page=page, search_query=search_query)

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
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
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

@app.route('/upload_resolution', methods=['POST'])
def upload_resolution_file():
    if 'file' not in request.files:
        return Response("No file part", status=400)

    files = request.files.getlist('file')
    for file in files:
        if file.filename == '':
            return Response("No selected file", status=400)

        if not file.filename.lower().endswith('.pdf'):
            return Response("Only PDF files are allowed", status=400)

        filename = secure_filename(file.filename)
        file_path = os.path.join(RESOLUTIONS_FOLDER, file.filename)
        file.save(file_path)
        text_chunks = extract_text_chunks(file_path)
        embeddings = generate_text_embeddings(text_chunks)
        if embeddings:
            embeddings_store[filename] = {'chunks': embeddings}
            save_json_store(JSON_STORE, embeddings_store)
    return jsonify({'answer': 'Resoluciones subidas y listas para procesar.'})

@app.route('/responses', methods=['GET'])
def get_saved_response():
    filename = request.args.get('filename', '')
    if filename in responses_store:
        return jsonify({'answer': responses_store[filename]})
    return jsonify({'answer': 'No se ha encontrado una respuesta almacenada.'})

@app.route('/delete', methods=['POST'])
def delete_file():
    data = request.json
    filename = data.get('filename', '')

    if filename in embeddings_store:
        del embeddings_store[filename]
        save_json_store(JSON_STORE, embeddings_store)

    if filename in responses_store:
        del responses_store[filename]
        save_json_store(RESPONSES_STORE, responses_store)

    # Optionally, delete the actual file from the uploads directory
    file_path = os.path.join(RESOLUTIONS_FOLDER, filename)
    if os.path.exists(file_path):
        os.remove(file_path)

    return jsonify({'answer': 'Resolución eliminada.'})

@app.route('/ask', methods=['POST'])
def ask_question():
    data = request.json
    filename = data.get('filename', '')
    if filename not in embeddings_store:
        return jsonify({'error': 'Documento no encontrado'}), 400
    document_data = embeddings_store[filename]
    context = "\n".join([chunk[0] for chunk in document_data['chunks']])
    prompt = f"Context:\n{context}\n\n{prompt_template}"
    response_text = generate_ai_response(prompt)
    responses_store[filename] = response_text
    save_json_store(RESPONSES_STORE, responses_store)
    return jsonify({'answer': response_text})

@app.route('/ask_all', methods=['POST'])
def ask_all():
    data = request.json
    filenames = data.get('filenames', [])
    responses = {}

    for filename in filenames:
        if filename in embeddings_store:
            document_data = embeddings_store[filename]
            context = "\n".join([chunk[0] for chunk in document_data['chunks']])
            prompt = f"Context:\n{context}\n\n{prompt_template}"
            response_text = generate_ai_response(prompt)
            responses[filename] = response_text
            responses_store[filename] = response_text
            save_json_store(RESPONSES_STORE, responses_store)

    return jsonify(responses)

@app.route('/export_csv', methods=['GET'])
def export_csv():
    csv_output = []
    responses_data = load_json_store(RESPONSES_STORE)

    if not responses_data:
        print("No hay datos en responses_store.json")
        return Response("No hay datos para exportar", mimetype="text/plain")

    for filename, response_text in responses_data.items():
        doc_name = filename.replace('.pdf', '').replace('.PDF', '')

        cleaned_text = re.sub(r'```json|```', '', response_text).strip()

        try:
            members_data = json.loads(cleaned_text)

            if not isinstance(members_data, list):
                print(f"Formato incorrecto en {filename}, se esperaba una lista.")
                continue

            for member in members_data:
                row = [
                    doc_name,
                    (member.get("Miembro del comité") or "").strip(),
                    (member.get("Cargo en la empresa") or "").strip(),
                    (member.get("Cargo en la comisión") or "").strip()
                ]
                csv_output.append(row)
        except json.JSONDecodeError as e:
            print(f"Error al decodificar JSON en {filename}: {e}")

    if not csv_output:
        print("No se extrajeron datos válidos del JSON.")
        return Response("No hay datos válidos para exportar", mimetype="text/plain")

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{timestamp}-CELEC-COM.csv"

    def generate():
        output = io.StringIO()
        output.write('\ufeff')
        writer = csv.writer(output, delimiter='\t')

        # Write headers
        writer.writerow(["Código del proceso", "Nombre del miembro", "Cargo en la empresa", "Cargo en la comisión"])
        for row in csv_output:
            writer.writerow(row)

        output.seek(0)
        yield output.read()
        output.close()

    return Response(generate(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename={filename}"})

@app.route('/delete_all', methods=['POST'])
def delete_all_file():
    # Make a copy of keys since we're modifying the dictionary during iteration
    filenames = list(embeddings_store.keys())
    
    for filename in filenames:
        # Remove from stores
        del embeddings_store[filename]
        if filename in responses_store:
            del responses_store[filename]
        
        # Remove physical file
        file_path = os.path.join(RESOLUTIONS_FOLDER, filename)
        if os.path.exists(file_path):
            os.remove(file_path)
    
    # Save the cleared stores
    save_json_store(JSON_STORE, embeddings_store)
    save_json_store(RESPONSES_STORE, responses_store)

    return jsonify({'message': 'Todas las resoluciones han sido eliminadas.'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
