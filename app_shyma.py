import os
from flask import Flask, render_template, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd

app = Flask(__name__)

# 1. Configuración de la Base de Datos
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///ast.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# 2. Modelo de Datos (La estructura de la tabla)
class RegistroAST(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    responsable = db.Column(db.String(100), nullable=False)
    area = db.Column(db.String(100), nullable=False)
    tarea = db.Column(db.Text, nullable=False)
    riesgo = db.Column(db.String(50), nullable=False)
    control = db.Column(db.String(100), nullable=False)
    # NUEVA COLUMNA: Se llena sola al crear el registro
    fecha = db.Column(db.DateTime, default=datetime.now)

# Crear la base de datos si no existe
with app.app_context():
    db.create_all()

# --- RUTAS DE LA APLICACIÓN ---

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/ast', methods=['POST'])
def crear_ast():
    nuevo_ast = RegistroAST(
        responsable = request.form.get('responsable'),
        area = request.form.get('area'),
        tarea = request.form.get('tarea'),
        riesgo = request.form.get('riesgo'),
        control = request.form.get('control')
    )
    db.session.add(nuevo_ast)
    db.session.commit()
    return redirect(url_for('ver_registros'))

@app.route('/registros')
def ver_registros():
    # Ordenar por fecha de forma descendente (los últimos primero)
    todos_los_ast = RegistroAST.query.order_by(RegistroAST.fecha.desc()).all()
    return render_template('ver_registros.html', registros=todos_los_ast)

@app.route('/eliminar/<int:id>')
def eliminar_ast(id):
    registro_a_borrar = RegistroAST.query.get_or_404(id)
    db.session.delete(registro_a_borrar)
    db.session.commit()
    return redirect(url_for('ver_registros'))

@app.route('/exportar')
def exportar_excel():
    registros = RegistroAST.query.all()
    datos = []
    for r in registros:
        datos.append({
            "FECHA Y HORA": r.fecha.strftime("%d/%m/%Y %H:%M"),
            "ID": r.id,
            "RESPONSABLE": r.responsable,
            "ÁREA": r.area,
            "TAREA": r.tarea,
            "RIESGO": r.riesgo,
            "MEDIDA DE CONTROL": r.control
        })
    
    df = pd.DataFrame(datos)
    nombre_archivo = f"Reporte_AST_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
    
    # Creamos el archivo usando un motor de escritura profesional
    with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Registros_Seguridad')
        
        # Accedemos a la hoja para darle formato
        workbook = writer.book
        ws = writer.sheets['Registros_Seguridad']
        
        # 1. Definir Estilos (Colores y Letras)
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid") # Azul oscuro
        header_font = Font(color="FFFFFF", bold=True, size=12) # Letra blanca negrita
        center_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                        top=Side(style='thin'), bottom=Side(style='thin'))

        # 2. Aplicar formato a la cabecera (la primera fila)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = border

        # 3. Auto-ajustar el ancho de las columnas según el contenido
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 4)
            ws.column_dimensions[column_letter].width = adjusted_width

        # 4. Mejoras de usabilidad
        ws.auto_filter.ref = ws.dimensions # Añade filtros automáticos
        ws.freeze_panes = "A2" # Congela la fila superior para que no se pierda al bajar

    return send_file(nombre_archivo, as_attachment=True)-ap

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
    
