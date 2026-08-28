from flask import Flask, render_template, request, jsonify, send_file
import os
import json
import pandas as pd
from datetime import datetime
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
audit_sessions = {}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_documents():
    data = request.json
    mock_facts = [
        {'id': 1, 'description': 'Cantidad de registros analizados en SAP', 'value': '15,432', 'source': 'SAP', 'reference': 'Reporte ZFI_AP_001', 'status': 'pending'},
        {'id': 2, 'description': 'Proveedores con diferencias en conciliaciones', 'value': '23', 'source': 'Excel', 'reference': 'Conciliaciones_Proveedores.xlsx', 'status': 'pending'},
        {'id': 3, 'description': 'Importe total de diferencias detectadas', 'value': '$3,175,700', 'source': 'Excel', 'reference': 'Resumen_Diferencias.xlsx', 'status': 'pending'},
        {'id': 4, 'description': 'Porcentaje de recepcion por debajo del 75% de vida util', 'value': '22%', 'source': 'SAP', 'reference': 'Reporte ZMM_INGRESOS', 'status': 'pending'},
        {'id': 5, 'description': 'Tickets Jira abiertos relacionados con el proceso', 'value': '7', 'source': 'Jira', 'reference': 'Project: AUDIT-2024', 'status': 'pending'},
        {'id': 6, 'description': 'Controles existentes documentados en procedimiento', 'value': '5', 'source': 'Procedimiento interno', 'reference': 'PROC-LOG-003 v2.1', 'status': 'pending'}
    ]
    return jsonify({'facts': mock_facts, 'message': 'Hechos extraidos. Revise y valide cada uno.'})

@app.route('/generate-memo', methods=['POST'])
def generate_memo():
    data = request.json
    audit_data = data.get('auditData', {})
    validated_facts = data.get('validatedFacts', [])
    tasks = data.get('tasks', [])
    results = data.get('results', [])
    findings = data.get('findings', [])
    style = data.get('style', 'ejecutivo')

    memo = {
        'header': {
            'titulo': audit_data.get('titulo', 'Auditoria'),
            'analisis': audit_data.get('analisis', ''),
            'sector': audit_data.get('sector', ''),
            'proceso': audit_data.get('proceso', ''),
            'periodo': audit_data.get('periodo', ''),
            'alcance': audit_data.get('alcance', ''),
            'auditor': audit_data.get('auditor', ''),
            'fecha': audit_data.get('fecha', datetime.now().strftime('%d/%m/%Y'))
        },
        'objetivos': audit_data.get('objetivos', []),
        'fuentes': audit_data.get('fuentes', []),
        'hechos_validados': validated_facts,
        'tareas': tasks,
        'resultados': results,
        'hallazgos': findings,
        'conclusiones': f"Se identificaron {len(findings)} hallazgo(s). Se recomienda implementar las acciones de mejora propuestas."
    }
    return jsonify({'memo': memo, 'message': 'Memo generado exitosamente'})

@app.route('/export-excel', methods=['POST'])
def export_excel():
    data = request.json
    memo_data = data.get('memo', {})
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        memo_df = pd.DataFrame([
            {'Campo': 'TITULO', 'Valor': memo_data.get('header', {}).get('titulo', '')},
            {'Campo': 'ANALISIS', 'Valor': memo_data.get('header', {}).get('analisis', '')},
            {'Campo': 'SECTOR', 'Valor': memo_data.get('header', {}).get('sector', '')},
            {'Campo': 'PERIODO', 'Valor': memo_data.get('header', {}).get('periodo', '')},
            {'Campo': 'ALCANCE', 'Valor': memo_data.get('header', {}).get('alcance', '')},
            {'Campo': 'AUDITOR', 'Valor': memo_data.get('header', {}).get('auditor', '')},
            {'Campo': 'FECHA', 'Valor': memo_data.get('header', {}).get('fecha', '')}
        ])
        memo_df.to_excel(writer, sheet_name='MEMO', index=False)

        hallazgos = memo_data.get('hallazgos', [])
        if hallazgos:
            hallazgos_df = pd.DataFrame([{
                'Titulo': h.get('titulo', ''), 'Criticidad': h.get('criticidad', ''),
                'Area_Responsable': h.get('area_responsable', ''), 'Riesgo': h.get('riesgo', ''),
                'Estado': h.get('estado', ''), 'Fecha_Objetivo': h.get('fecha_objetivo', '')
            } for h in hallazgos])
            hallazgos_df.to_excel(writer, sheet_name='Hallazgos', index=False)
        else:
            pd.DataFrame({'Mensaje': ['No hay hallazgos']}).to_excel(writer, sheet_name='Hallazgos', index=False)

        resultados = memo_data.get('resultados', [])
        if resultados:
            pd.DataFrame(resultados).to_excel(writer, sheet_name='Resultados', index=False)
        else:
            pd.DataFrame({'Mensaje': ['No hay resultados']}).to_excel(writer, sheet_name='Resultados', index=False)

        fuentes = memo_data.get('fuentes', [])
        if fuentes:
            pd.DataFrame([{'Fuente': f} for f in fuentes]).to_excel(writer, sheet_name='Fuentes', index=False)
        else:
            pd.DataFrame({'Mensaje': ['No hay fuentes']}).to_excel(writer, sheet_name='Fuentes', index=False)

        hechos = memo_data.get('hechos_validados', [])
        if hechos:
            traz_df = pd.DataFrame([{
                'Descripcion': h.get('description', ''), 'Valor': h.get('value', ''),
                'Fuente': h.get('source', ''), 'Referencia': h.get('reference', '')
            } for h in hechos])
            traz_df.to_excel(writer, sheet_name='Trazabilidad', index=False)
        else:
            pd.DataFrame({'Mensaje': ['No hay hechos']}).to_excel(writer, sheet_name='Trazabilidad', index=False)

    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'audit_memo_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')

@app.route('/improve-text', methods=['POST'])
def improve_text():
    data = request.json
    return jsonify({'original': data.get('text', ''), 'improved': data.get('text', ''), 'message': 'Redaccion revisada'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
