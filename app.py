import os
import io
import base64
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docxtpl import DocxTemplate

app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get('SECRET_KEY', 'cambia_esto_por_un_clave_segura')

# — Usuarios y permisos —
USERS = {
    os.environ.get('BASIC_USER',   'admin'):    os.environ.get('BASIC_PASS',   'password'),
    os.environ.get('BASIC_USER2', 'usuario2'): os.environ.get('BASIC_PASS2', 'pass2')
}
PERMISSIONS = {
    'usuario2': [
        'ITINERARIO','AMENIDAD','PRE ARRIVAL NOTAS (BORRADOR)',
        'REGISTRO DE CONTACTO','Transfer'
    ]
}

def check_auth(u, p):
    return USERS.get(u) == p

def authenticate():
    return ('Autorización requerida.'), 401, {
        'WWW-Authenticate': 'Basic realm="Login Required"'
    }

def requires_auth(f):
    @wraps(f)
    def deco(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return deco

# — Mapeo de nombres para la UI —
display_names = {
    'Pms_Confirm_No':   'Número de confirmación',
    'CSV_Guest_NM':     'Nombre huésped',
    'CSV_Cust_Email':   'Email',
    'CSV_Arrival_Date': 'Check-in',
    'Which_Date':       'Fecha reserva',
    # ... añade más si las necesitas ...
}

# — Configuración de Google Sheets —
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly"
]
if 'GOOGLE_SHEETS_JSON_B64' in os.environ:
    raw = base64.b64decode(os.environ['GOOGLE_SHEETS_JSON_B64'])
    CRED_FILE = '/tmp/credentials.json'
    with open(CRED_FILE, 'wb') as f: f.write(raw)
else:
    BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
    CRED_FILE = os.path.join(BASE_DIR, 'credentials.json')

creds     = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, SCOPES)
client    = gspread.authorize(creds)
SHEET_ID  = '1LDhajDpQTzi0RLw8BXLTzmA1m9yRlTX_SrxC9aKLKYg'
worksheet = client.open_by_key(SHEET_ID).worksheet('hoja')

@app.route('/', methods=['GET'])
@requires_auth
def index():
    return render_template('index.html', display_names=display_names)

@app.route('/search', methods=['POST'])
@requires_auth
def search():
    auth    = request.authorization
    user    = auth.username
    allowed = PERMISSIONS.get(user)

    # Campos que el usuario puede usar para filtrar
    fields = ['Pms_Confirm_No','CSV_Guest_NM','CSV_Cust_Email',
              'CSV_Arrival_Date','Which_Date']

    # Recogemos sólo los campos que no estén vacíos
    criteria = {
        f: request.form.get(f).strip()
        for f in fields
        if request.form.get(f) and request.form.get(f).strip()
    }
    if not criteria:
        flash('Debes indicar al menos un criterio de búsqueda.', 'error')
        return redirect(url_for('index'))

    # Leemos la primera fila (headers) y creamos un map de header->col_index
    headers     = worksheet.row_values(1)
    col_indices = {h: i+1 for i, h in enumerate(headers)}

    # Para cada criterio, buscamos filas que coincidan exacto (lower-case vs lower-case)
    row_sets = []
    for field, val in criteria.items():
        col = col_indices.get(field)
        if not col:
            row_sets.append(set())
            continue

        # obtener todos los valores de la columna
        col_values = worksheet.col_values(col)  # incluye header en index 0
        matching = {
            idx
            for idx, cell in enumerate(col_values[1:], start=2)
            if cell.strip().lower() == val.lower()
        }
        row_sets.append(matching)

    # Intersección de todos los sets (AND)
    matches = set.intersection(*row_sets) if row_sets else set()

    if not matches:
        flash('No se encontraron reservas que cumplan todos los criterios.', 'error')
        return redirect(url_for('index'))

    # Si sólo hay una fila, vamos directo a editar
    if len(matches) == 1:
        row = matches.pop()
        vals   = worksheet.row_values(row)
        record = dict(zip(headers, vals))
        record['row_idx'] = row
        return render_template('edit.html',
                               record=record,
                               headers=headers,
                               user=user,
                               allowed=allowed,
                               display_names=display_names)

    # Si hay varias, listamos sólo los campos clave con botón Editar
    records = []
    for row in sorted(matches):
        vals = worksheet.row_values(row)
        rec  = dict(zip(headers, vals))
        records.append({
            'CSV_Guest_NM':   rec.get('CSV_Guest_NM',''),
            'CSV_Cust_Email': rec.get('CSV_Cust_Email',''),
            'Pms_Confirm_No': rec.get('Pms_Confirm_No',''),
            'row_idx':        row
        })

    return render_template('index.html',
                           records=records,
                           display_names=display_names)

@app.route('/edit/<int:row_idx>', methods=['GET'])
@requires_auth
def edit_record(row_idx):
    auth    = request.authorization
    user    = auth.username
    allowed = PERMISSIONS.get(user)
    headers = worksheet.row_values(1)
    vals    = worksheet.row_values(row_idx)
    record  = dict(zip(headers, vals))
    record['row_idx'] = row_idx

    return render_template('edit.html',
                           record=record,
                           headers=headers,
                           user=user,
                           allowed=allowed,
                           display_names=display_names)

@app.route('/update', methods=['POST'])
@requires_auth
def update():
    auth    = request.authorization
    user    = auth.username
    allowed = PERMISSIONS.get(user)
    row_idx = int(request.form.get('row_idx'))
    headers = worksheet.row_values(1)

    for i, h in enumerate(headers, start=1):
        if allowed and h not in allowed:
            continue
        worksheet.update_cell(row_idx, i, request.form.get(h, ''))

    if request.form.get('export'):
        context = {
            h.lower().replace(' ','_'): request.form.get(h,'')
            for h in headers
        }
        tpl = os.path.join(os.path.dirname(__file__),
                           'templates_docx','itinerary_template.docx')
        doc = DocxTemplate(tpl)
        doc.render(context)

        bio = io.BytesIO()
        doc.save(bio); bio.seek(0)
        fname = f"Itinerary_{context.get('pms_confirm_no', row_idx)}.docx"
        return send_file(bio,
                         as_attachment=True,
                         download_name=fname,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    flash('Registro actualizado con éxito.', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
