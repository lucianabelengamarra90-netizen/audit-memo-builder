from flask import Flask, render_template, request, jsonify, send_file

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from docx import Document
from pypdf import PdfReader
from openai import OpenAI

from io import BytesIO, StringIO
from tempfile import NamedTemporaryFile
from datetime import datetime

import csv
import os
import re
import sqlite3
import unicodedata
import uuid
import zipfile
import xml.etree.ElementTree as ET


# ============================================================
# APP
# ============================================================

app = Flask(__name__)

app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024


# ============================================================
# CONFIGURACIÓN
# ============================================================

ALLOWED_EXTENSIONS = {
    "xlsx",
    "csv",
    "docx",
    "pdf",
    "txt",
}

# Máximo que mostramos en pantalla por solapa.
# IMPORTANTE:
# esto NO corta el análisis.
MAX_DISPLAY_ITEMS_PER_SHEET = 250

# Longitud máxima del texto mostrado por fila.
MAX_ROW_TEXT_LENGTH = 8000


# ============================================================
# PALABRAS CLAVE DE AUDITORÍA
# ============================================================

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
        "prueba realizada",
        "pruebas realizadas",
        "recalculo",
        "recálculo",
        "cruce realizado",
        "cruce",
        "validacion",
        "validación",
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
        "plan de acción",
    ],

    "Comentario": [
        "comentario",
        "comentarios",
    ],
}


# ============================================================
# HELPERS DE TEXTO
# ============================================================

def normalize_text(value):

    if value is None:
        return ""

    text = str(value).strip()

    if not text:
        return ""

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


def contains_letters(text):

    if not text:
        return False

    return any(
        char.isalpha()
        for char in text
    )


def row_to_text(values):

    parts = []

    for value in values:

        text = clean_text(value)

        if text:
            parts.append(text)

    result = " | ".join(parts)

    if len(result) > MAX_ROW_TEXT_LENGTH:
        result = result[:MAX_ROW_TEXT_LENGTH]

    return result


# ============================================================
# PALABRAS CLAVE NORMALIZADAS
#
# Se calculan UNA SOLA VEZ.
# ============================================================

NORMALIZED_KEYWORDS = {}

for category, keywords in KEYWORD_GROUPS.items():

    NORMALIZED_KEYWORDS[category] = []

    for keyword in keywords:

        NORMALIZED_KEYWORDS[
            category
        ].append(
            (
                normalize_text(keyword),
                keyword
            )
        )


def detect_categories(text):

    normalized = normalize_text(text)

    if not normalized:
        return []

    matches = []

    for category, keyword_pairs in NORMALIZED_KEYWORDS.items():

        for normalized_keyword, original_keyword in keyword_pairs:

            if (
                normalized_keyword
                and normalized_keyword in normalized
            ):

                matches.append({
                    "category": category,
                    "keyword": original_keyword
                })

                break

    return matches


# ============================================================
# ITEMS
# ============================================================

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


def add_unique_item(
    items,
    seen,
    item
):

    fingerprint = (
        item["filename"],
        item["originName"],
        item["reference"],
        item["category"],
        normalize_text(
            item["text"]
        )
    )

    if fingerprint in seen:
        return False

    seen.add(
        fingerprint
    )

    items.append(
        item
    )

    return True


# ============================================================
# OPENAI
# ============================================================

def get_openai_client():

    api_key = os.getenv(
        "OPENAI_API_KEY",
        ""
    ).strip()

    if not api_key:
        return None

    return OpenAI(
        api_key=api_key
    )


# ============================================================
# XLSX
# ============================================================

MAIN_NS = (
    "http://schemas.openxmlformats.org/"
    "spreadsheetml/2006/main"
)

REL_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)


def qname(namespace, tag):

    return f"{{{namespace}}}{tag}"


# ============================================================
# SHARED STRINGS
#
# Guardamos:
# - texto
# - si contiene keyword
# - categorías detectadas
#
# De esta forma NO analizamos nuevamente cada texto
# cada vez que aparece en una hoja.
# ============================================================

def create_shared_strings_db(
    archive,
    db_path
):

    connection = sqlite3.connect(
        db_path
    )

    cursor = connection.cursor()

    cursor.execute(
        "PRAGMA journal_mode=OFF"
    )

    cursor.execute(
        "PRAGMA synchronous=OFF"
    )

    cursor.execute(
        "PRAGMA temp_store=MEMORY"
    )

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS shared_strings (
            idx INTEGER PRIMARY KEY,
            value TEXT,
            relevant INTEGER DEFAULT 0,
            categories TEXT DEFAULT '',
            keywords TEXT DEFAULT ''
        )
        """
    )

    if "xl/sharedStrings.xml" not in archive.namelist():

        connection.commit()

        return connection

    print(
        "Analizando sharedStrings...",
        flush=True
    )

    index = 0

    batch = []

    with archive.open(
        "xl/sharedStrings.xml"
    ) as stream:

        context = ET.iterparse(
            stream,
            events=("end",)
        )

        for _, element in context:

            if element.tag != qname(
                MAIN_NS,
                "si"
            ):
                continue

            parts = []

            for text_element in element.iter(
                qname(
                    MAIN_NS,
                    "t"
                )
            ):

                if text_element.text:
                    parts.append(
                        text_element.text
                    )

            value = "".join(
                parts
            )

            matches = detect_categories(
                value
            )

            relevant = 1 if matches else 0

            categories = "|||".join(
                match["category"]
                for match in matches
            )

            keywords = "|||".join(
                match["keyword"]
                for match in matches
            )

            batch.append(
                (
                    index,
                    value,
                    relevant,
                    categories,
                    keywords
                )
            )

            index += 1

            if len(batch) >= 10000:

                cursor.executemany(
                    """
                    INSERT INTO shared_strings (
                        idx,
                        value,
                        relevant,
                        categories,
                        keywords
                    )
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    batch
                )

                connection.commit()

                batch.clear()

            element.clear()

    if batch:

        cursor.executemany(
            """
            INSERT INTO shared_strings (
                idx,
                value,
                relevant,
                categories,
                keywords
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            batch
        )

    connection.commit()

    cursor.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_shared_relevant
        ON shared_strings(relevant)
        """
    )

    connection.commit()

    print(
        f"SharedStrings analizados: {index}",
        flush=True
    )

    return connection


# ============================================================
# SOLAPAS DEL XLSX
# ============================================================

def get_workbook_sheet_paths(
    archive
):

    workbook_root = ET.fromstring(
        archive.read(
            "xl/workbook.xml"
        )
    )

    relationships_root = ET.fromstring(
        archive.read(
            "xl/_rels/workbook.xml.rels"
        )
    )

    relationships = {}

    for relationship in relationships_root:

        relationship_id = relationship.attrib.get(
            "Id"
        )

        target = relationship.attrib.get(
            "Target",
            ""
        )

        if not target:
            continue

        if target.startswith("/"):
            target = target.lstrip("/")

        elif not target.startswith("xl/"):
            target = "xl/" + target

        target = os.path.normpath(
            target
        ).replace(
            "\\",
            "/"
        )

        relationships[
            relationship_id
        ] = target

    sheets = []

    sheets_element = workbook_root.find(
        qname(
            MAIN_NS,
            "sheets"
        )
    )

    if sheets_element is None:
        return sheets

    archive_names = set(
        archive.namelist()
    )

    for sheet in sheets_element:

        sheet_name = sheet.attrib.get(
            "name",
            "Sin nombre"
        )

        relationship_id = sheet.attrib.get(
            qname(
                REL_NS,
                "id"
            )
        )

        sheet_path = relationships.get(
            relationship_id
        )

        if (
            sheet_path
            and sheet_path in archive_names
        ):

            sheets.append({
                "name": sheet_name,
                "path": sheet_path,
                "state": sheet.attrib.get(
                    "state",
                    "visible"
                )
            })

    return sheets


# ============================================================
# SHARED STRING LOOKUP
# ============================================================

def get_shared_string_info(
    connection,
    index
):

    row = connection.execute(
        """
        SELECT
            value,
            relevant,
            categories,
            keywords
        FROM shared_strings
        WHERE idx = ?
        """,
        (index,)
    ).fetchone()

    if not row:

        return {
            "value": "",
            "relevant": False,
            "matches": []
        }

    value = row[0] or ""

    relevant = bool(
        row[1]
    )

    categories = (
        row[2].split("|||")
        if row[2]
        else []
    )

    keywords = (
        row[3].split("|||")
        if row[3]
        else []
    )

    matches = []

    for category, keyword in zip(
        categories,
        keywords
    ):

        matches.append({
            "category": category,
            "keyword": keyword
        })

    return {
        "value": value,
        "relevant": relevant,
        "matches": matches
    }


# ============================================================
# CELDAS XLSX
# ============================================================

def extract_inline_string(
    cell_element
):

    inline_string = cell_element.find(
        qname(
            MAIN_NS,
            "is"
        )
    )

    if inline_string is None:
        return ""

    parts = []

    for text_element in inline_string.iter(
        qname(
            MAIN_NS,
            "t"
        )
    ):

        if text_element.text:
            parts.append(
                text_element.text
            )

    return "".join(
        parts
    )


def get_raw_cell_value(
    cell_element
):

    value_element = cell_element.find(
        qname(
            MAIN_NS,
            "v"
        )
    )

    if (
        value_element is None
        or value_element.text is None
    ):
        return ""

    return value_element.text


# ============================================================
# EXTRAER XLSX
# ============================================================

def extract_xlsx(
    uploaded_file,
    filename
):

    items = []

    seen = set()

    warnings = []

    excel_path = None

    db_path = None

    shared_connection = None

    try:

        # ----------------------------------------------------
        # Guardar archivo temporalmente en disco
        # ----------------------------------------------------

        with NamedTemporaryFile(
            suffix=".xlsx",
            delete=False
        ) as temp_excel:

            excel_path = temp_excel.name

            while True:

                chunk = uploaded_file.stream.read(
                    1024 * 1024
                )

                if not chunk:
                    break

                temp_excel.write(
                    chunk
                )

        # ----------------------------------------------------
        # SQLite temporal
        # ----------------------------------------------------

        with NamedTemporaryFile(
            suffix=".sqlite",
            delete=False
        ) as temp_db:

            db_path = temp_db.name

        # ----------------------------------------------------
        # Abrir XLSX
        # ----------------------------------------------------

        with zipfile.ZipFile(
            excel_path,
            "r"
        ) as archive:

            shared_connection = create_shared_strings_db(
                archive,
                db_path
            )

            sheets = get_workbook_sheet_paths(
                archive
            )

            print(
                f"{filename}: {len(sheets)} solapas detectadas.",
                flush=True
            )

            # ------------------------------------------------
            # Recorrer TODAS las solapas
            # ------------------------------------------------

            for sheet_index, sheet in enumerate(
                sheets,
                start=1
            ):

                sheet_name = sheet["name"]

                sheet_path = sheet["path"]

                print(
                    (
                        f"[{sheet_index}/{len(sheets)}] "
                        f"Procesando solapa: {sheet_name}"
                    ),
                    flush=True
                )

                rows_processed = 0

                relevant_rows = 0

                displayed_items = 0

                omitted_items = 0

                # --------------------------------------------
                # La solapa se recorre COMPLETA
                # --------------------------------------------

                with archive.open(
                    sheet_path
                ) as stream:

                    context = ET.iterparse(
                        stream,
                        events=("end",)
                    )

                    for _, element in context:

                        if element.tag != qname(
                            MAIN_NS,
                            "row"
                        ):
                            continue

                        rows_processed += 1

                        row_number = element.attrib.get(
                            "r",
                            ""
                        )

                        row_values = []

                        row_matches = []

                        row_has_candidate = False

                        # ------------------------------------
                        # Leer las celdas de la fila
                        # ------------------------------------

                        for cell in element.findall(
                            qname(
                                MAIN_NS,
                                "c"
                            )
                        ):

                            cell_type = cell.attrib.get(
                                "t"
                            )

                            # --------------------------------
                            # SHARED STRING
                            # --------------------------------

                            if cell_type == "s":

                                raw_value = get_raw_cell_value(
                                    cell
                                )

                                if not raw_value:
                                    continue

                                try:

                                    shared_index = int(
                                        raw_value
                                    )

                                except Exception:

                                    continue

                                info = get_shared_string_info(
                                    shared_connection,
                                    shared_index
                                )

                                value = info["value"]

                                if value:
                                    row_values.append(
                                        value
                                    )

                                if info["relevant"]:

                                    row_has_candidate = True

                                    row_matches.extend(
                                        info["matches"]
                                    )

                            # --------------------------------
                            # INLINE STRING
                            # --------------------------------

                            elif cell_type == "inlineStr":

                                value = extract_inline_string(
                                    cell
                                )

                                if value:
                                    row_values.append(
                                        value
                                    )

                                matches = detect_categories(
                                    value
                                )

                                if matches:

                                    row_has_candidate = True

                                    row_matches.extend(
                                        matches
                                    )

                            # --------------------------------
                            # STR
                            # --------------------------------

                            elif cell_type == "str":

                                value = get_raw_cell_value(
                                    cell
                                )

                                if value:
                                    row_values.append(
                                        value
                                    )

                                matches = detect_categories(
                                    value
                                )

                                if matches:

                                    row_has_candidate = True

                                    row_matches.extend(
                                        matches
                                    )

                            # --------------------------------
                            # NUMÉRICOS / FECHAS / BOOLEANOS
                            # Se guardan para reconstruir contexto
                            # solamente.
                            # --------------------------------

                            else:

                                value = get_raw_cell_value(
                                    cell
                                )

                                if value:
                                    row_values.append(
                                        value
                                    )

                        # ------------------------------------
                        # Si no hay candidato, seguimos
                        # ------------------------------------

                        if not row_has_candidate:

                            element.clear()

                            continue

                        # ------------------------------------
                        # Reconstrucción completa de fila
                        # ------------------------------------

                        row_text = row_to_text(
                            row_values
                        )

                        if not row_text:

                            element.clear()

                            continue

                        relevant_rows += 1

                        # ------------------------------------
                        # Eliminar categorías duplicadas
                        # dentro de la misma fila
                        # ------------------------------------

                        unique_matches = []

                        seen_categories = set()

                        for match in row_matches:

                            key = (
                                match["category"],
                                match["keyword"]
                            )

                            if key in seen_categories:
                                continue

                            seen_categories.add(
                                key
                            )

                            unique_matches.append(
                                match
                            )

                        # ------------------------------------
                        # Guardar para frontend
                        # ------------------------------------

                        for match in unique_matches:

                            if (
                                displayed_items
                                >= MAX_DISPLAY_ITEMS_PER_SHEET
                            ):

                                omitted_items += 1

                                continue

                            item = make_item(
                                category=match[
                                    "category"
                                ],
                                text=row_text,
                                filename=filename,
                                origin_type="Excel",
                                origin_name=sheet_name,
                                reference=(
                                    f"Fila {row_number}"
                                    if row_number
                                    else ""
                                ),
                                keyword=match[
                                    "keyword"
                                ]
                            )

                            added = add_unique_item(
                                items,
                                seen,
                                item
                            )

                            if added:

                                displayed_items += 1

                        element.clear()

                # --------------------------------------------
                # Log de avance
                # --------------------------------------------

                print(
                    (
                        f"Solapa finalizada: {sheet_name} | "
                        f"Filas: {rows_processed} | "
                        f"Filas relevantes: {relevant_rows} | "
                        f"Mostradas: {displayed_items}"
                    ),
                    flush=True
                )

                if omitted_items > 0:

                    warnings.append(
                        (
                            f"{filename} · {sheet_name}: "
                            f"se detectaron más elementos de los "
                            f"{MAX_DISPLAY_ITEMS_PER_SHEET} mostrados. "
                            "El análisis recorrió la solapa completa, "
                            "pero se limitó la cantidad visualizada "
                            "para evitar sobrecargar el navegador."
                        )
                    )

    except zipfile.BadZipFile:

        raise ValueError(
            "El archivo Excel no es válido o está dañado."
        )

    finally:

        if shared_connection:

            try:
                shared_connection.close()
            except Exception:
                pass

        for path in (
            excel_path,
            db_path
        ):

            if not path:
                continue

            try:
                os.remove(
                    path
                )
            except Exception:
                pass

    return items, warnings


# ============================================================
# CSV
# ============================================================

def extract_csv(
    uploaded_file,
    filename
):

    items = []

    seen = set()

    raw = uploaded_file.read()

    try:

        text = raw.decode(
            "utf-8-sig"
        )

    except Exception:

        text = raw.decode(
            "latin1",
            errors="replace"
        )

    stream = StringIO(
        text
    )

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

        row_text = row_to_text(
            row
        )

        if not row_text:
            continue

        if not contains_letters(
            row_text
        ):
            continue

        matches = detect_categories(
            row_text
        )

        for match in matches:

            item = make_item(
                category=match[
                    "category"
                ],
                text=row_text,
                filename=filename,
                origin_type="CSV",
                origin_name="CSV",
                reference=f"Fila {row_number}",
                keyword=match[
                    "keyword"
                ]
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

def extract_docx(
    uploaded_file,
    filename
):

    items = []

    seen = set()

    document = Document(
        BytesIO(
            uploaded_file.read()
        )
    )

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
                category=match[
                    "category"
                ],
                text=text,
                filename=filename,
                origin_type="Word",
                origin_name="Documento",
                reference=(
                    f"Párrafo {paragraph_number}"
                ),
                keyword=match[
                    "keyword"
                ]
            )

            add_unique_item(
                items,
                seen,
                item
            )

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
                    category=match[
                        "category"
                    ],
                    text=row_text,
                    filename=filename,
                    origin_type="Word",
                    origin_name=(
                        f"Tabla {table_number}"
                    ),
                    reference=(
                        f"Fila {row_number}"
                    ),
                    keyword=match[
                        "keyword"
                    ]
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

def extract_pdf(
    uploaded_file,
    filename
):

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

            page_text = (
                page.extract_text()
                or ""
            )

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

        for index, line in enumerate(
            lines
        ):

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
                    category=match[
                        "category"
                    ],
                    text=context,
                    filename=filename,
                    origin_type="PDF",
                    origin_name=(
                        f"Página {page_number}"
                    ),
                    reference=(
                        f"Página {page_number}"
                    ),
                    keyword=match[
                        "keyword"
                    ]
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
                origin_name="Documento"
            )
        )

    return items


# ============================================================
# TXT
# ============================================================

def extract_txt(
    uploaded_file,
    filename
):

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
                category=match[
                    "category"
                ],
                text=line,
                filename=filename,
                origin_type="TXT",
                origin_name="Documento",
                reference=(
                    f"Línea {line_number}"
                ),
                keyword=match[
                    "keyword"
                ]
            )

            add_unique_item(
                items,
                seen,
                item
            )

    return items


# ============================================================
# TEXTO MANUAL
# ============================================================

def extract_free_text(
    text
):

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

        for match in matches:

            item = make_item(
                category=match[
                    "category"
                ],
                text=line,
                filename="Texto ingresado",
                origin_type="Texto",
                origin_name="Ingreso manual",
                reference=(
                    f"Línea {line_number}"
                ),
                keyword=match[
                    "keyword"
                ]
            )

            add_unique_item(
                items,
                seen,
                item
            )

    return items


# ============================================================
# HOME
# ============================================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


# ============================================================
# HEALTH
# ============================================================

@app.route("/health")
def health():

    openai_configured = bool(
        os.getenv(
            "OPENAI_API_KEY",
            ""
        ).strip()
    )

    return jsonify({
        "status": "ok",
        "openaiConfigured": openai_configured
    })


# ============================================================
# EXTRAER INFORMACIÓN
# ============================================================

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

    warnings = []

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

        print(
            f"Iniciando extracción: {filename}",
            flush=True
        )

        try:

            if extension == "xlsx":

                (
                    excel_items,
                    excel_warnings
                ) = extract_xlsx(
                    uploaded_file,
                    filename
                )

                extracted_items.extend(
                    excel_items
                )

                warnings.extend(
                    excel_warnings
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

        except Exception as exc:

            print(
                (
                    f"Error procesando {filename}: "
                    f"{type(exc).__name__}: {exc}"
                ),
                flush=True
            )

            errors.append(
                (
                    f"{filename}: "
                    f"{type(exc).__name__}: {exc}"
                )
            )

        print(
            f"Extracción finalizada: {filename}",
            flush=True
        )

    if free_text:

        extracted_items.extend(
            extract_free_text(
                free_text
            )
        )

    return jsonify({
        "items": extracted_items,
        "count": len(
            extracted_items
        ),
        "errors": errors,
        "warnings": warnings,
        "message": (
            f"Se identificaron "
            f"{len(extracted_items)} "
            "elementos potencialmente relevantes. "
            "Revisalos antes de incorporarlos al memo."
        )
    })


# Compatibilidad con versión anterior
@app.route(
    "/analyze",
    methods=["POST"]
)
def analyze_alias():

    return extract_information()


# ============================================================
# IA
# ============================================================

@app.route(
    "/improve-text",
    methods=["POST"]
)
def improve_text():

    data = request.get_json(
        silent=True
    ) or {}

    text = clean_text(
        data.get(
            "text",
            ""
        )
    )

    field_type = clean_text(
        data.get(
            "fieldType",
            "Texto de auditoría"
        )
    )

    if not text:

        return jsonify({
            "error": (
                "Primero escribí una idea "
                "o texto para mejorar."
            )
        }), 400

    openai_client = get_openai_client()

    if not openai_client:

        return jsonify({
            "error": (
                "Render no está entregando "
                "OPENAI_API_KEY al proceso."
            )
        }), 500

    instructions = """
Actuá como especialista en redacción de Auditoría Interna,
con experiencia en procesos corporativos y retail mayorista.

Tu función es mejorar exclusivamente la redacción aportada
por el auditor.

No realices la auditoría.
No inventes hechos ni evidencia.

REGLAS:

- No inventes importes.
- No inventes cantidades.
- No inventes fechas.
- No inventes períodos.
- No inventes proveedores.
- No inventes clientes.
- No inventes personas.
- No inventes sistemas.
- No inventes documentos.
- No inventes controles.
- No inventes muestras.
- No inventes resultados.
- No inventes hallazgos.
- No inventes evidencia.
- No inventes conclusiones.
- No atribuyas causas no informadas.

Conservá el significado original.

Utilizá lenguaje claro, técnico, profesional y corporativo.

OBJETIVO:
Mejorá y completá conceptualmente el objetivo sin inventar
procedimientos específicos.

ALCANCE:
Ordená únicamente la información suministrada.

SITUACIÓN OBSERVADA:
Describí objetivamente los hechos aportados.

RIESGO:
Mejorá la formulación del riesgo sin inventar impactos concretos.

PROPUESTA DE MEJORA:
Profesionalizá la recomendación orientándola a fortalecer
controles, trazabilidad y eficiencia sin modificar
la intención original.

Devolvé solamente el texto final.
"""

    prompt = f"""
TIPO DE CAMPO:
{field_type}

TEXTO DEL AUDITOR:
{text}

Mejorá la redacción respetando estrictamente las reglas.
"""

    try:

        response = openai_client.responses.create(
            model="gpt-5.6-luna",
            instructions=instructions,
            input=prompt
        )

        improved = clean_text(
            response.output_text
        )

        if not improved:

            return jsonify({
                "error": (
                    "La IA no devolvió una redacción."
                )
            }), 500

        return jsonify({
            "original": text,
            "improved": improved
        })

    except Exception as exc:

        print(
            (
                "Error OpenAI /improve-text: "
                f"{type(exc).__name__}: {exc}"
            ),
            flush=True
        )

        return jsonify({
            "error": (
                "La solicitud a OpenAI no pudo completarse. "
                "Revisá el log de Render."
            )
        }), 500


# ============================================================
# EXPORTACIÓN
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


def auto_width(
    worksheet
):

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
        (
            "Auditoría",
            general.get(
                "title",
                ""
            )
        ),
        (
            "Área",
            general.get(
                "area",
                ""
            )
        ),
        (
            "Proceso",
            general.get(
                "process",
                ""
            )
        ),
        (
            "Período",
            general.get(
                "period",
                ""
            )
        ),
        (
            "Auditor",
            general.get(
                "auditor",
                ""
            )
        ),
        (
            "Objetivo",
            general.get(
                "objective",
                ""
            )
        ),
        (
            "Alcance",
            general.get(
                "scope",
                ""
            )
        ),
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

    row += 2

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

    auto_width(
        ws
    )

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

        format_header(
            cell
        )

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

    auto_width(
        ws_findings
    )

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

        format_header(
            cell
        )

    for source in sources:

        ws_sources.append([
            source.get("name", ""),
            source.get("type", ""),
            source.get("reference", ""),
            source.get("description", ""),
        ])

    auto_width(
        ws_sources
    )

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
        "Incluido"
    ])

    for cell in ws_trace[1]:

        format_header(
            cell
        )

    for item in extracted:

        ws_trace.append([
            item.get("category", ""),
            item.get("text", ""),
            item.get("filename", ""),
            item.get("originName", ""),
            item.get("reference", ""),
            item.get("keyword", ""),
            (
                "Sí"
                if item.get("included")
                else "No"
            ),
        ])

    auto_width(
        ws_trace
    )

    # ========================================================
    # DESCARGA
    # ========================================================

    output = BytesIO()

    workbook.save(
        output
    )

    output.seek(
        0
    )

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
