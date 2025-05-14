from flask import Flask, render_template, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)
# --- Cambia user:pass@host/db por tus credenciales MySQL ---
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://user:pass@host/db'
db = SQLAlchemy(app)

class Item(db.Model):
    __tablename__ = 'tu_tabla'
    id = db.Column(db.Integer, primary_key=True)
    columna_a = db.Column(db.String(100))
    columna_b = db.Column(db.String(100))
    # …añade aquí todas tus columnas…

@app.route('/')
def form():
    return render_template('edit.html')

@app.route('/api/row/<int:id>', methods=['GET'])
def get_row(id):
    row = Item.query.get_or_404(id)
    return jsonify({c.name: getattr(row, c.name) for c in row.__table__.columns})

@app.route('/api/row/<int:id>', methods=['POST'])
def update_row(id):
    data = request.json
    row = Item.query.get_or_404(id)
    for k, v in data.items():
        if hasattr(row, k):
            setattr(row, k, v)
    db.session.commit()
    return jsonify(success=True)

@app.route('/api/row/<int:id>/pdf')
def make_pdf(id):
    row = Item.query.get_or_404(id)
    buf = BytesIO()
    p = canvas.Canvas(buf, pagesize=A4)
    text = p.beginText(50, 800)
    for c in row.__table__.columns:
        text.textLine(f"{c.name}: {getattr(row, c.name)}")
    p.drawText(text)
    p.showPage(); p.save()
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"row_{id}.pdf",
                     mimetype='application/pdf')

if __name__ == '__main__':
    app.run(debug=True)
