import os
import json
import io
import base64
import unicodedata
import re

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from functools import wraps
from docxtpl import DocxTemplate

# —– CONFIGURACIÓN DE FLASK —–
app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get('SECRET_KEY', 'cambia_esto_por_un_valor_seguro')

# —– AUTENTICACIÓN BÁSICA —–
def check_auth(username, password):
    return username == os.environ.get('BASIC_USER', 'admin') \
       and password == os.environ.get('BASIC_PASS', 'password')

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

# —– SLUGIFY PARA ENCABEZADOS —–
def slugify(value: str) -> str:
    text = unicodedata.normalize('NFKD', value).encode('ascii','ignore').decode('ascii')
    text = re.sub(r'[^\w\s-]', '', text).strip().lower()
    return re.sub(r'[-\s]+', '_', text)

# —– GOOGLE SHEETS SETUP —–
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly"
]

# credenciales JSON en Base64 (Elastic Beanstalk env var)
if 'GOOGLE_SHEETS_JSON_B64' in os.environ:
    raw = base64.b64decode(os.environ['GOOGLE_SHEETS_JSON_B64'])
    CRED_FILE = '/tmp/credentials.json'
    with open(CRED_FILE, 'wb') as f:
        f.write(raw)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    CRED_FILE = os.path.join(BASE_DIR, 'credentials.json')

creds  = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, SCOPES)
client = gspread.authorize(creds)

# ID y worksheet
SHEET_ID  = '1LDhajDpQTzi0RLw8BXLTzmA1m9yRlTX_SrxC9aKLKYg'
ws        = client.open_by_key(SHEET_ID).worksheet('hoja')

# Leemos encabezados y generamos slugs
raw_headers = ws.row_values(1)
slugged     = [slugify(h) for h in raw_headers]
header_map  = dict(zip(raw_headers, slugged))
# Ejemplo: header_map["Viaja con Mascota"] == "viaja_con_mascota"

# —– RUTAS —–
@app.route('/', methods=['GET','POST'])
def index():
    if request.method == 'POST':
        search_id = request.form.get('search_id','').strip()
        if not search_id:
            flash('El campo ID no puede estar vacío.','error')
        else:
            try:
                cell = ws.find(search_id, in_column=3)
                row  = cell.row
                values = ws.row_values(row)
                record = dict(zip(raw_headers, values))
                record['row_idx'] = row
                return render_template('edit.html', record=record)
            except Exception:
                flash('ID no encontrado. Try again.','error')
    return render_template('index.html')

@app.route('/update', methods=['POST'])
@requires_auth
def update():
    row_idx = int(request.form.get('row_idx'))
    # 1) Guardar cambios (permite vacíos)
    try:
        for col, h in enumerate(raw_headers, start=1):
            ws.update_cell(row_idx, col, request.form.get(h,''))
    except Exception:
        flash('Error al guardar cambios. Intenta de nuevo.','error')
        record = {h: request.form.get(h,'') for h in raw_headers}
        record['row_idx'] = row_idx
        return render_template('edit.html', record=record)

    # 2) Si piden exportar Word, preparamos contexto con slugs
    if request.form.get('export'):
        values  = [request.form.get(h,'') for h in raw_headers]
        context = { header_map[h]: v for h,v in zip(raw_headers, values) }

        tpl = DocxTemplate('templates_docx/itinerary_template.docx')
        tpl.render(context)

        bio = io.BytesIO()
        tpl.save(bio)
        bio.seek(0)
        fn = f"Itinerary_{context.get('pms_confirm_no','')}.docx"
        return send_file(
            bio,
            as_attachment=True,
            download_name=fn,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    flash('Registro actualizado con éxito.','success')
    return redirect(url_for('index'))

if __name__=='__main__':
    app.run(debug=True)
