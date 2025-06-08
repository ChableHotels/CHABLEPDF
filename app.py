import os
import json
import io
import base64
import unicodedata

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from functools import wraps
from docxtpl import DocxTemplate

# —– Configuración de Flask —–
app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get('SECRET_KEY', 'cambia_esto_por_un_clave_segura')

# —– Usuarios y permisos —–
USERS = {
    os.environ.get('BASIC_USER',   'admin'):    os.environ.get('BASIC_PASS',   'password'),
    os.environ.get('BASIC_USER2', 'usuario2'): os.environ.get('BASIC_PASS2', 'pass2')
}
PERMISSIONS = {
    'usuario2': [
        'ITINERARIO',
        'AMENIDAD',
        'PRE ARRIVAL NOTAS (BORRADOR)',
        'REGISTRO DE CONTACTO',
        'Transfer'
    ]
}

def check_auth(username, password):
    return USERS.get(username) == password

def authenticate():
    return ('Autorización requerida.'), 401, {
        'WWW-Authenticate': 'Basic realm="Login Required"'
    }

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# —– Mapeo de nombres para mostrar —–
display_names = {
    'Hotel_Long_Nm': 'Hotel',
    'Which_Date': 'Fecha reserva',
    'Pms_Confirm_No': 'Número de confirmación',
    'CSV_Guest_NM': 'Nombre huésped',
    'CSV_Cust_Phone1': 'Teléfono',
    'CSV_Cust_Email': 'Email',
    'CSV_Arrival_Date': 'Check-in',
    'CSV_Depart_Date': 'Check-out',
    'CSV_Nights_Qty': 'Cantidad de noches',
    'CSV_Status': 'Estatus reserva',
    'cuantosAcomp': 'Número de acompañantes',
    'detallesAcompanantes': 'Detalle de acompañantes',
    'Hora de Llegada': 'Hora de llegada',
    'AM/PM Llegada': 'AM/PM llegada',
    'Viaja con Mascota': 'Viaja con mascota',
    'Detalles Mascota': 'Detalle mascota',
    'Motivo del viaje': 'Motivo de viaje',
    'Origen Vistita': 'Nombre visita',
    'Restricciones alimenticias': 'Restricciones alimenticias',
    'Detalles Alergia': 'Detalle alergias',
    'Bebidas y Platillos Preferidos': 'Bebidas y platillos',
    'Preferencia Café y Lácteos': 'Preferencia café y lácteos',
    'Hobbies Favoritos': 'Hobbies favoritos',
    'Solicitud Especial': 'Solicitud especial',
    'Comentarios Adicionales': 'Comentarios adicionales',
    'CASITA': 'Casita',
    'TIPO DE HUÉSPED': 'Tipo de huésped',
    'RESERVA': 'Reserva',
    'AMENIDAD/CELEBRACIÓN': 'Amenidad/Celebración',
    'ITINERARIO': 'Itinerario',
    'AMENIDAD': 'Amenidad',
    'PRE ARRIVAL NOTAS (BORRADOR)': 'Pre–arrival notas',
    'REGISTRO DE CONTACTO': 'Registro de contacto',
    'Transfer': 'Transfer',
    'Aerolinea': 'Aerolínea',
    'Numero de vuelo': 'Número de vuelo',
    'Aeropuerto de origen': 'Aeropuerto origen',
    'Aeropuerto destino (salida)': 'Aeropuerto destino',
    'horario del vuelo': 'Horario de vuelo',
}

# —– Google Sheets setup —–
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly"
]

if 'GOOGLE_SHEETS_JSON_B64' in os.environ:
    raw = base64.b64decode(os.environ['GOOGLE_SHEETS_JSON_B64'])
    CRED_FILE = '/tmp/credentials.json'
    with open(CRED_FILE, 'wb') as f:
        f.write(raw)
else:
    BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
    CRED_FILE = os.path.join(BASE_DIR, 'credentials.json')

creds  = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, SCOPES)
client = gspread.authorize(creds)

SHEET_ID   = '1LDhajDpQTzi0RLw8BXLTzmA1m9yRlTX_SrxC9aKLKYg'
worksheet  = client.open_by_key(SHEET_ID).worksheet('hoja')

# —– Función auxiliar para normalizar —–
def normalize_string(s: str) -> str:
    import re, unicodedata
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = re.sub(r'[^A-Za-z0-9]', '_', s).lower()
    s = re.sub(r'__+', '_', s).strip('_')
    return s

# —– Rutas —–
@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    auth = request.authorization
    user = auth.username if auth else None
    allowed = PERMISSIONS.get(user)

    search_id = request.form.get('search_id','').strip()
    if not search_id:
        flash('El campo ID no puede estar vacío.', 'error')
        return redirect(url_for('index'))

    try:
        cell    = worksheet.find(search_id, in_column=3)
        row_idx = cell.row
        headers = worksheet.row_values(1)
        values  = worksheet.row_values(row_idx)
        record  = dict(zip(headers, values))
        record['row_idx'] = row_idx

        return render_template(
            'edit.html',
            record=record,
            headers=headers,
            user=user,
            allowed=allowed,
            display_names=display_names
        )
    except Exception:
        flash('ID no encontrado.', 'error')
        return redirect(url_for('index'))

@app.route('/update', methods=['POST'])
@requires_auth
def update():
    auth    = request.authorization
    user    = auth.username
    allowed = PERMISSIONS.get(user)
    row_idx = int(request.form.get('row_idx'))
    headers = worksheet.row_values(1)

    # Guardar en Sheets
    for i, h in enumerate(headers, start=1):
        if allowed is not None and h not in allowed:
            continue
        worksheet.update_cell(row_idx, i, request.form.get(h, ''))

    # Exportar Word
    if request.form.get('export'):
        context = {}
        for h in headers:
            key_n = normalize_string(h)
            context[key_n] = request.form.get(h, '')

        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        tpl_path = os.path.join(BASE_DIR, 'templates_docx', 'itinerary_template.docx')
        doc = DocxTemplate(tpl_path)
        doc.render(context)

        bio = io.BytesIO()
        doc.save(bio); bio.seek(0)
        fname = f"Itinerary_{context.get('pms_confirm_no', row_idx)}.docx"

        return send_file(
            bio,
            as_attachment=True,
            download_name=fname,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    flash('Registro actualizado con éxito.', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
