from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from pypdf import PdfReader
from io import BytesIO, StringIO
from tempfile import NamedTemporaryFile
from datetime import datetime
import csv
import os
import re
import unicodedata
import uuid

app = Flask(__name__)

app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024


# ============================================================
# CONFIGURACIÓN
# ============================================================

ALLOWED_EXTENSIONS = {
    "xlsx",
    "xls",
    "csv",
    "docx",
    "pdf",
    "txt",
}


KEYWORD_GROUPS = {
    "Hallazgo": [
        "hallazgo",
        "hallazgos",
    ],
    "Conclusión": [
        "conclusion",
        "conclusiones",
    ],
    "Diferencia": [
        "diferencia",
        "diferencias",
        "discrepancia",
        "discrepancias",
    ],
    "Observación": [
        "observacion",
        "observaciones",
        "desvio",
        "desvios",
    ],
    "Resultado": [
        "resultado",
        "resultados",
        "sin excepcion",
        "sin diferencias",
        "sin observaciones",
    ],
    "Riesgo": [
        "riesgo",
        "riesgos",
    ],
    "Propuesta de mejora": [
        "propuesta",
        "propuesta de mejora",
        "recomendacion",
        "recomendaciones",
        "mejora",
    ],
    "Objetivo": [
        "objetivo",
        "objetivos",
    ],
    "Alcance": [
        "alcance",
    ],
    "Tarea realizada": [
        "tarea",
        "tareas",
        "tarea realizada",
        "tareas realizadas",
        "procedimiento realizado",
        "procedimientos realizados",
        "trabajo realizado",
    ],
    "Incumplimiento": [
        "incumplimiento",
        "incumplimientos",
    ],
    "Pendiente": [
        "pendiente",
        "pendientes",
    ],
    "Acción": [
        "accion",
        "acciones",
        "plan de accion",
    ],
    "Comentario": [
        "comentario",
        "comentarios",
    ],
}


# ============================================================
# HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip()

    text = unicodedata.normalize(
        "NFD",
        text
    )

    text = "".join(
        char
        for char in text
        if unicodedata.category(char) != "Mn"
    )

    text = text.lower()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def clean_text(value):
    if value is None:
        return ""

    text = str(value)

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def get_extension(filename):
    if "." not in filename:
        return ""

    return filename.rsplit(
        ".",
        1
    )[-1].lower()


def detect_categories(text):
    normalized = normalize_text(text)

    matches = []

    for category, keywords in KEYWORD_GROUPS.items():
        for keyword in keywords:

            normalized_keyword = normalize_text(
                keyword
            )

            if normalized_keyword in normalized:

                matches.append({
                    "category": category,
                    "keyword": keyword
                })

                break

    return matches


def make_item(
    category,
    text,
    filename,
    origin_type,
    origin_name="",
    reference="",
    keyword=""
):
    return {
        "id": str(uuid.uuid4()),
        "category": category,
        "text": clean_text(text),
        "filename": filename,
        "originType": origin_type,
        "originName": origin_name,
        "reference": reference,
        "keyword": keyword,
        "included": False,
        "converted": False,
    }


def add_unique_item(items, seen, item):
    fingerprint = (
        item["filename"],
        item["originName"],
        item["category"],
        normalize_text(item["text"])
    )

    if fingerprint in seen:
        return

    seen.add(fingerprint)
    items.append(item)


def row_to_text(values):
    parts = []

    for value in values:
        text = clean_text(value)

        if text:
            parts.append(text)

    return " | ".join(parts)


# ============================================================
# EXCEL XLSX
# ============================================================

def extract_xlsx(uploaded_file, filename):
    items = []
    seen = set()

    with NamedTemporaryFile(
        suffix=".xlsx",
        delete=False
    ) as temp_file:

        temp_path = temp_file.name

        while True:
            chunk = uploaded_file.stream.read(
                1024 * 1024
            )

            if not chunk:
                break

            temp_file.write(chunk)

    workbook = None

    try:
        workbook = load_workbook(
            temp_path,
            read_only=True,
            data_only=True
        )

        for worksheet in workbook.worksheets:

            sheet_name = worksheet.title

            # ------------------------------------------------
            # También revisamos el nombre de la solapa
            # ------------------------------------------------

            sheet_matches = detect_categories(
                sheet_name
            )

            rows_detected = 0

            for row_number, row in enumerate(
                worksheet.iter_rows(
                    values_only=True
                ),
                start=1
            ):

                row_text = row_to_text(row)

                if not row_text:
                    continue

                matches = detect_categories(
                    row_text
                )

                if not matches and sheet_matches:
                    matches = sheet_matches

                if not matches:
                    continue

                for match in matches:

                    item = make_item(
                        category=match["category"],
                        text=row_text,
                        filename=filename,
                        origin_type="Excel",
                        origin_name=sheet_name,
                        reference=f"Fila {row_number}",
                        keyword=match["keyword"]
                    )

                    add_unique_item(
                        items,
                        seen,
                        item
                    )

                rows_detected += 1

                # Evita respuestas inmanejables cuando una
                # solapa completa está titulada "Hallazgos".
                if rows_detected >= 150:
                    break

    finally:

        if workbook:
            workbook.close()

        try:
            os.remove(temp_path)
        except Exception:
            pass

    return items


# ============================================================
# CSV
# ============================================================

def extract_csv(uploaded_file, filename):
    items = []
    seen = set()

    raw = uploaded_file.read()

    encoding = "utf-8-sig"

    try:
        text = raw.decode(
            encoding
        )
    except Exception:
        text = raw.decode(
            "latin1",
            errors="replace"
        )

    stream = StringIO(text)

    sample = text[:5000]

    try:
        dialect = csv.Sniffer().sniff(
            sample,
            delimiters=",;\t|"
        )
    except Exception:
        dialect = csv.excel

    reader = csv.reader(
        stream,
        dialect
    )

    for row_number, row in enumerate(
        reader,
        start=1
    ):

        row_text = row_to_text(row)

        if not row_text:
            continue

        matches = detect_categories(
            row_text
        )

        for match in matches:

            item = make_item(
                category=match["category"],
                text=row_text,
                filename=filename,
                origin_type="CSV",
                origin_name="CSV",
                reference=f"Fila {row_number}",
                keyword=match["keyword"]
            )

            add_unique_item(
                items,
                seen,
                item
            )

    return items


# ============================================================
# WORD
# ============================================================

def extract_docx(uploaded_file, filename):
    items = []
    seen = set()

    document = Document(
        BytesIO(
            uploaded_file.read()
        )
    )

    # --------------------------------------------------------
    # PÁRRAFOS
    # --------------------------------------------------------

    for paragraph_number, paragraph in enumerate(
        document.paragraphs,
        start=1
    ):

        text = clean_text(
            paragraph.text
        )

        if not text:
            continue

        matches = detect_categories(
            text
        )

        for match in matches:

            item = make_item(
                category=match["category"],
                text=text,
                filename=filename,
                origin_type="Word",
                origin_name="Documento",
                reference=f"Párrafo {paragraph_number}",
                keyword=match["keyword"]
            )

            add_unique_item(
                items,
                seen,
                item
            )

    # --------------------------------------------------------
    # TABLAS
    # --------------------------------------------------------

    for table_number, table in enumerate(
        document.tables,
        start=1
    ):

        for row_number, row in enumerate(
            table.rows,
            start=1
        ):

            values = [
                cell.text
                for cell in row.cells
            ]

            row_text = row_to_text(
                values
            )

            if not row_text:
                continue

            matches = detect_categories(
                row_text
            )

            for match in matches:

                item = make_item(
                    category=match["category"],
                    text=row_text,
                    filename=filename,
                    origin_type="Word",
                    origin_name=f"Tabla {table_number}",
                    reference=f"Fila {row_number}",
                    keyword=match["keyword"]
                )

                add_unique_item(
                    items,
                    seen,
                    item
                )

    return items


# ============================================================
# PDF
# ============================================================

def extract_pdf(uploaded_file, filename):
    items = []
    seen = set()

    reader = PdfReader(
        BytesIO(
            uploaded_file.read()
        )
    )

    pages_with_text = 0

    for page_number, page in enumerate(
        reader.pages,
        start=1
    ):

        try:
            page_text = page.extract_text() or ""
        except Exception:
            page_text = ""

        if not page_text.strip():
            continue

        pages_with_text += 1

        lines = [
            clean_text(line)
            for line in page_text.splitlines()
            if clean_text(line)
        ]

        # ----------------------------------------------------
        # Revisamos bloques de líneas para conservar contexto
        # ----------------------------------------------------

        for index, line in enumerate(lines):

            matches = detect_categories(
                line
            )

            if not matches:
                continue

            start = max(
                0,
                index - 1
            )

            end = min(
                len(lines),
                index + 3
            )

            context = " ".join(
                lines[start:end]
            )

            for match in matches:

                item = make_item(
                    category=match["category"],
                    text=context,
                    filename=filename,
                    origin_type="PDF",
                    origin_name=f"Página {page_number}",
                    reference=f"Página {page_number}",
                    keyword=match["keyword"]
                )

                add_unique_item(
                    items,
                    seen,
                    item
                )

    if pages_with_text == 0:

        items.append(
            make_item(
                category="Advertencia",
                text=(
                    "No se pudo extraer texto del PDF. "
                    "El documento puede estar escaneado."
                ),
                filename=filename,
                origin_type="PDF",
                origin_name="Documento",
                reference="",
                keyword=""
            )
        )

    return items


# ============================================================
# TXT
# ============================================================

def extract_txt(uploaded_file, filename):
    items = []
    seen = set()

    raw = uploaded_file.read()

    text = None

    for encoding in (
        "utf-8-sig",
        "utf-8",
        "latin1"
    ):

        try:
            text = raw.decode(
                encoding
            )
            break
        except Exception:
            pass

    if text is None:
        return items

    lines = [
        clean_text(line)
        for line in text.splitlines()
        if clean_text(line)
    ]

    for line_number, line in enumerate(
        lines,
        start=1
    ):

        matches = detect_categories(
            line
        )

        for match in matches:

            item = make_item(
                category=match["category"],
                text=line,
                filename=filename,
                origin_type="TXT",
                origin_name="Documento",
                reference=f"Línea {line_number}",
                keyword=match["keyword"]
            )

            add_unique_item(
                items,
                seen,
                item
            )

    return items


# ============================================================
# FREE TEXT
# ============================================================

def extract_free_text(text):
    items = []
    seen = set()

    lines = [
        clean_text(line)
        for line in text.splitlines()
        if clean_text(line)
    ]

    for line_number, line in enumerate(
        lines,
        start=1
    ):

        matches = detect_categories(
            line
        )

        if not matches:
            continue

        for match in matches:

            item = make_item(
                category=match["category"],
                text=line,
                filename="Texto ingresado",
                origin_type="Texto",
                origin_name="Ingreso manual",
                reference=f"Línea {line_number}",
                keyword=match["keyword"]
            )

            add_unique_item(
                items,
                seen,
                item
            )

    return items


# ============================================================
# ROUTES
# ============================================================

@app.route("/")
def index():
    return render_template(
        "index.html"
    )


@app.route(
    "/extract",
    methods=["POST"]
)
def extract_information():

    files = request.files.getlist(
        "files"
    )

    free_text = request.form.get(
        "freeText",
        ""
    ).strip()

    if not files and not free_text:

        return jsonify({
            "error": (
                "Cargá al menos un papel de trabajo "
                "o ingresá texto adicional."
            )
        }), 400

    extracted_items = []
    errors = []

    for uploaded_file in files:

        filename = (
            uploaded_file.filename
            or "archivo_sin_nombre"
        )

        extension = get_extension(
            filename
        )

        if extension not in ALLOWED_EXTENSIONS:

            errors.append(
                f"{filename}: formato no admitido."
            )

            continue

        try:

            if extension == "xlsx":

                extracted_items.extend(
                    extract_xlsx(
                        uploaded_file,
                        filename
                    )
                )

            elif extension == "csv":

                extracted_items.extend(
                    extract_csv(
                        uploaded_file,
                        filename
                    )
                )

            elif extension == "docx":

                extracted_items.extend(
                    extract_docx(
                        uploaded_file,
                        filename
                    )
                )

            elif extension == "pdf":

                extracted_items.extend(
                    extract_pdf(
                        uploaded_file,
                        filename
                    )
                )

            elif extension == "txt":

                extracted_items.extend(
                    extract_txt(
                        uploaded_file,
                        filename
                    )
                )

            elif extension == "xls":

                errors.append(
                    f"{filename}: el formato XLS antiguo "
                    "debe convertirse a XLSX para conservar "
                    "la lectura eficiente por solapas."
                )

        except Exception as exc:

            errors.append(
                f"{filename}: {str(exc)}"
            )

    if free_text:

        extracted_items.extend(
            extract_free_text(
                free_text
            )
        )

    return jsonify({
        "items": extracted_items,
        "count": len(extracted_items),
        "errors": errors,
        "message": (
            f"Se identificaron {len(extracted_items)} "
            "elementos potencialmente relevantes. "
            "Revisalos antes de incorporarlos al memo."
        )
    })


# Alias para evitar problemas si queda alguna llamada vieja.
@app.route(
    "/analyze",
    methods=["POST"]
)
def analyze_alias():
    return extract_information()


# ============================================================
# EXPORTAR EXCEL
# ============================================================

def format_header(cell):
    cell.fill = PatternFill(
        "solid",
        fgColor="17365D"
    )

    cell.font = Font(
        color="FFFFFF",
        bold=True
    )

    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )


def auto_width(worksheet):
    for column in worksheet.columns:

        max_length = 0

        column_letter = get_column_letter(
            column[0].column
        )

        for cell in column:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

            if cell.value is None:
                continue

            length = len(
                str(cell.value)
            )

            max_length = max(
                max_length,
                min(
                    length,
                    70
                )
            )

        worksheet.column_dimensions[
            column_letter
        ].width = max(
            12,
            min(
                max_length + 2,
                70
            )
        )


@app.route(
    "/export-excel",
    methods=["POST"]
)
def export_excel():

    payload = request.get_json(
        silent=True
    ) or {}

    memo = payload.get(
        "memo",
        {}
    )

    general = memo.get(
        "general",
        {}
    )

    findings = memo.get(
        "findings",
        []
    )

    sources = memo.get(
        "sources",
        []
    )

    extracted = memo.get(
        "extracted",
        []
    )

    workbook = Workbook()

    # ========================================================
    # MEMO
    # ========================================================

    ws = workbook.active
    ws.title = "Memo"

    ws.merge_cells(
        "A1:F1"
    )

    ws["A1"] = "AUDITORÍA INTERNA"

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="17365D"
    )

    ws["A1"].font = Font(
        color="FFFFFF",
        bold=True,
        size=14
    )

    ws["A1"].alignment = Alignment(
        horizontal="center"
    )

    row = 3

    general_rows = [
        ("Auditoría", general.get("title", "")),
        ("Área", general.get("area", "")),
        ("Proceso", general.get("process", "")),
        ("Período", general.get("period", "")),
        ("Auditor", general.get("auditor", "")),
        ("Objetivo", general.get("objective", "")),
        ("Alcance", general.get("scope", "")),
    ]

    for label, value in general_rows:

        ws.cell(
            row=row,
            column=1,
            value=label
        )

        ws.cell(
            row=row,
            column=1
        ).font = Font(
            bold=True,
            color="17365D"
        )

        ws.merge_cells(
            start_row=row,
            start_column=2,
            end_row=row,
            end_column=6
        )

        ws.cell(
            row=row,
            column=2,
            value=value
        )

        row += 1

    row += 1

    for index, finding in enumerate(
        findings,
        start=1
    ):

        ws.merge_cells(
            start_row=row,
            start_column=1,
            end_row=row,
            end_column=6
        )

        ws.cell(
            row=row,
            column=1,
            value=(
                f"HALLAZGO {index:02d} - "
                f"{finding.get('title', '')}"
            )
        )

        ws.cell(
            row=row,
            column=1
        ).fill = PatternFill(
            "solid",
            fgColor="EAF2F8"
        )

        ws.cell(
            row=row,
            column=1
        ).font = Font(
            bold=True,
            color="17365D"
        )

        row += 1

        details = [
            (
                "Situación observada",
                finding.get(
                    "situation",
                    ""
                )
            ),
            (
                "Riesgo",
                finding.get(
                    "risk",
                    ""
                )
            ),
            (
                "Propuesta de mejora",
                finding.get(
                    "proposal",
                    ""
                )
            ),
            (
                "Área responsable",
                finding.get(
                    "responsibleArea",
                    ""
                )
            ),
            (
                "Criticidad",
                finding.get(
                    "severity",
                    ""
                )
            ),
            (
                "Estado",
                finding.get(
                    "status",
                    ""
                )
            ),
            (
                "Fecha compromiso",
                finding.get(
                    "targetDate",
                    ""
                )
            ),
            (
                "Archivo de origen",
                finding.get(
                    "sourceFile",
                    ""
                )
            ),
            (
                "Solapa / origen",
                finding.get(
                    "sourceLocation",
                    ""
                )
            ),
        ]

        for label, value in details:

            ws.cell(
                row=row,
                column=1,
                value=label
            )

            ws.cell(
                row=row,
                column=1
            ).font = Font(
                bold=True
            )

            ws.merge_cells(
                start_row=row,
                start_column=2,
                end_row=row,
                end_column=6
            )

            ws.cell(
                row=row,
                column=2,
                value=value
            )

            row += 1

        row += 2

    auto_width(ws)

    # ========================================================
    # HALLAZGOS
    # ========================================================

    ws_findings = workbook.create_sheet(
        "Hallazgos"
    )

    finding_headers = [
        "N°",
        "Título",
        "Situación observada",
        "Riesgo",
        "Propuesta de mejora",
        "Área responsable",
        "Responsable plan",
        "Criticidad",
        "Estado",
        "Fecha compromiso",
        "Base cuantitativa",
        "Archivo origen",
        "Solapa / origen",
        "Evidencia / referencia",
        "Ticket",
        "Seguimiento",
    ]

    ws_findings.append(
        finding_headers
    )

    for cell in ws_findings[1]:
        format_header(cell)

    for index, finding in enumerate(
        findings,
        start=1
    ):

        ws_findings.append([
            index,
            finding.get("title", ""),
            finding.get("situation", ""),
            finding.get("risk", ""),
            finding.get("proposal", ""),
            finding.get("responsibleArea", ""),
            finding.get("actionOwner", ""),
            finding.get("severity", ""),
            finding.get("status", ""),
            finding.get("targetDate", ""),
            finding.get("quantitativeBasis", ""),
            finding.get("sourceFile", ""),
            finding.get("sourceLocation", ""),
            finding.get("evidence", ""),
            finding.get("ticket", ""),
            finding.get("followUp", ""),
        ])

    ws_findings.freeze_panes = "A2"
    auto_width(ws_findings)

    # ========================================================
    # FUENTES
    # ========================================================

    ws_sources = workbook.create_sheet(
        "Fuentes"
    )

    ws_sources.append([
        "Nombre",
        "Tipo",
        "Referencia",
        "Descripción"
    ])

    for cell in ws_sources[1]:
        format_header(cell)

    for source in sources:

        ws_sources.append([
            source.get("name", ""),
            source.get("type", ""),
            source.get("reference", ""),
            source.get("description", ""),
        ])

    auto_width(ws_sources)

    # ========================================================
    # TRAZABILIDAD
    # ========================================================

    ws_trace = workbook.create_sheet(
        "Trazabilidad"
    )

    ws_trace.append([
        "Tipo detectado",
        "Contenido",
        "Archivo",
        "Origen / solapa",
        "Referencia",
        "Palabra clave",
        "Incluido",
    ])

    for cell in ws_trace[1]:
        format_header(cell)

    for item in extracted:

        ws_trace.append([
            item.get("category", ""),
            item.get("text", ""),
            item.get("filename", ""),
            item.get("originName", ""),
            item.get("reference", ""),
            item.get("keyword", ""),
            "Sí" if item.get("included") else "No",
        ])

    auto_width(ws_trace)

    # ========================================================
    # GUARDAR
    # ========================================================

    output = BytesIO()

    workbook.save(
        output
    )

    output.seek(0)

    filename = (
        "Audit_Memo_"
        + datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + ".xlsx"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
