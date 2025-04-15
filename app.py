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

from Services.ExcelService import export_excel_process_inf, export_excel_process_pro, export_excel_process_reg, export_excel_process_pre, load_hyperlinks_from_excel, save_dataframe_to_excel, export_technical_commission_to_excel, save_dataframe_prov_to_excel
from Services.LlmService import configure_gemini_api, generate_text_embeddings, generate_ai_response
from Services.PdfService import allowed_file, extract_text_chunks
from Services.ScraperService import ScraperService, fetch_admin, fetch_provider
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

@app.route('/providers')
def providers():
    delete_files_in_directory(OUTPUT_FOLDER)
    delete_files_in_directory(UPLOAD_FOLDER)
    return render_template('providers.html')

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

@app.route('/upload_prov', methods=['POST'])
def upload_prov_file():
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
    data = fetch_provider(modified_links)

    # Create DataFrame and save to Excel
    df = pd.DataFrame(data)
    output_file_path = save_dataframe_prov_to_excel(df, file.filename)

    return send_file(output_file_path, as_attachment=True)

@app.route('/upload_resolution', methods=['POST'])
def upload_resolution_file():
    try:
        if 'file' not in request.files:
            return jsonify({'message': 'No se ha seleccionado ningún archivo'}), 400

        files = request.files.getlist('file')
        processed_files = []
        
        for file in files:
            if file.filename == '':
                continue

            if not file.filename.lower().endswith('.pdf'):
                return jsonify({'message': 'Solo se permiten archivos PDF'}), 400

            filename = secure_filename(file.filename)
            file_path = os.path.join(RESOLUTIONS_FOLDER, file.filename)
            file.save(file_path)
            
            text_chunks = extract_text_chunks(file_path)
            embeddings = generate_text_embeddings(text_chunks)
            
            if embeddings:
                embeddings_store[filename] = {'chunks': embeddings}
                processed_files.append(filename)
        
        if processed_files:
            save_json_store(JSON_STORE, embeddings_store)
            return jsonify({'message': f'Se han procesado {len(processed_files)} archivos exitosamente'})
        else:
            return jsonify({'message': 'No se pudo procesar ningún archivo'}), 400
            
    except Exception as e:
        return jsonify({'message': f'Error al procesar los archivos: {str(e)}'}), 500

@app.route('/responses', methods=['GET'])
def get_saved_response():
    try:
        filename = request.args.get('filename')
        if not filename:
            return jsonify({'message': 'Nombre de archivo no proporcionado'}), 400
            
        response = responses_store.get(filename)
        if response:
            return jsonify({'answer': response})
        return jsonify({'message': 'No se ha encontrado una respuesta almacenada.'}), 404
    except Exception as e:
        return jsonify({'message': f'Error al obtener la respuesta: {str(e)}'}), 500

@app.route('/delete', methods=['POST'])
def delete_file():
    try:
        data = request.json
        if not data or 'filename' not in data:
            return jsonify({'message': 'Nombre de archivo no proporcionado'}), 400

        filename = data['filename']
        
        # Remove from stores
        embeddings_store.pop(filename, None)
        responses_store.pop(filename, None)
        
        # Save updated stores
        save_json_store(JSON_STORE, embeddings_store)
        save_json_store(RESPONSES_STORE, responses_store)

        # Remove physical file
        file_path = os.path.join(RESOLUTIONS_FOLDER, filename)
        if os.path.exists(file_path):
            os.remove(file_path)

        return jsonify({'message': 'Resolución eliminada exitosamente'})
    except Exception as e:
        return jsonify({'message': f'Error al eliminar el archivo: {str(e)}'}), 500

@app.route('/ask', methods=['POST'])
def ask_question():
    try:
        data = request.json
        if not data or 'filename' not in data:
            return jsonify({'message': 'Nombre de archivo no proporcionado'}), 400

        filename = data['filename']
        if filename not in embeddings_store:
            return jsonify({'message': 'Documento no encontrado'}), 404

        document_data = embeddings_store[filename]
        context = "\n".join([chunk[0] for chunk in document_data['chunks']])
        prompt = f"Context:\n{context}\n\n{prompt_template}"
        
        response_text = generate_ai_response(prompt)
        if not response_text:
            return jsonify({'message': 'No se pudo generar una respuesta'}), 500

        responses_store[filename] = response_text
        save_json_store(RESPONSES_STORE, responses_store)
        
        return jsonify({'answer': response_text})
    except Exception as e:
        return jsonify({'message': f'Error al procesar la solicitud: {str(e)}'}), 500

@app.route('/ask_all', methods=['POST'])
def ask_all():
    try:
        data = request.json
        if not data or 'filenames' not in data:
            return jsonify({'message': 'Lista de archivos no proporcionada'}), 400

        filenames = data['filenames']
        responses = {}
        errors = {}

        for filename in filenames:
            try:
                if filename not in embeddings_store:
                    errors[filename] = 'Documento no encontrado'
                    continue

                document_data = embeddings_store[filename]
                context = "\n".join([chunk[0] for chunk in document_data['chunks']])
                prompt = f"Context:\n{context}\n\n{prompt_template}"
                
                response_text = generate_ai_response(prompt)
                if response_text:
                    responses[filename] = response_text
                    responses_store[filename] = response_text
                else:
                    errors[filename] = 'No se pudo generar una respuesta'
            except Exception as e:
                errors[filename] = str(e)

        save_json_store(RESPONSES_STORE, responses_store)

        return jsonify({
            'responses': responses,
            'errors': errors,
            'message': f'Procesados {len(responses)} archivos, {len(errors)} errores'
        })
    except Exception as e:
        return jsonify({'message': f'Error al procesar la solicitud: {str(e)}'}), 500

@app.route('/export_csv', methods=['GET'])
def export_csv():
    try:
        responses_data = load_json_store(RESPONSES_STORE)

        if not responses_data:
            return jsonify({'message': 'No hay datos para exportar'}), 404

        output_file_path = export_technical_commission_to_excel(responses_data)
        
        return send_file(
            output_file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True
        )
    except Exception as e:
        return jsonify({'message': f'Error al exportar Excel: {str(e)}'}), 500

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
