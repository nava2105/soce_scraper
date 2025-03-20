from flask import Flask, request, Response, render_template, send_file
from Services.ScraperService import ScraperService
from Services.ExcelService import export_excel_process_inf, export_excel_process_pro, export_excel_process_reg, export_excel_process_pre
import os

scraper = ScraperService()

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

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


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)