from flask import Flask, render_template, request, jsonify, send_file
from datetime import datetime
from io import BytesIO
import re

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from pypdf import PdfReader


app = Flask(__name__)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


ALLOWED_EXTENSIONS = {
    "xlsx",
    "xls",
    "csv",
    "docx",
    "pdf",
    "txt"
}


MONEY_KEYWORDS = (
    "importe",
    "monto",
    "saldo",
    "total",
    "valor",
    "deuda",
    "capital",
    "pago",
    "cuota",
    "debe",
    "haber"
)


DATE_KEYWORDS = (
    "fecha",
    "date",
    "vto",
    "venc",
    "vencimiento",
    "emision",
    "emisión",
    "contabil"
)


# =========================================================
# AUXILIARES
# =========================================================

def clean_text(value):

    return re.sub(
        r"\s+",
        " ",
        str(value or "")
    ).strip()


def add_fact(
    facts,
    description,
    value,
    source,
    reference,
    kind="fact"
):

    facts.append({

        "id":
            len(facts) + 1,

        "description":
            clean_text(description),

        "value":
            clean_text(value),

        "source":
            clean_text(source),

        "reference":
            clean_text(reference),

        "status":
            "pending",

        "kind":
            kind
    })


# =========================================================
# CSV
# =========================================================

def read_csv_bytes(raw):

    last_error = None

    for encoding in (
        "utf-8-sig",
        "utf-8",
        "latin1"
    ):

        try:

            return pd.read_csv(

                BytesIO(raw),

                low_memory=False,

                encoding=encoding
            )

        except Exception as exc:

            last_error = exc


    raise last_error


# =========================================================
# IMPORTES
# =========================================================

def numeric_series(series):

    if pd.api.types.is_numeric_dtype(
        series
    ):

        return pd.to_numeric(
            series,
            errors="coerce"
        )


    text = (

        series
            .astype(str)
            .str.strip()
            .str.replace(
                r"[^0-9,.\-()]",
                "",
                regex=True
            )
            .str.replace(
                "(",
                "-",
                regex=False
            )
            .str.replace(
                ")",
                "",
                regex=False
            )
    )


    def parse_one(value):

        if value in (
            "",
            "-",
            ".",
            ",",
            "nan",
            "None"
        ):

            return None


        try:

            if (
                "," in value
                and
                "." in value
            ):

                if (
                    value.rfind(",")
                    >
                    value.rfind(".")
                ):

                    value = (
                        value
                        .replace(
                            ".",
                            ""
                        )
                        .replace(
                            ",",
                            "."
                        )
                    )

                else:

                    value = value.replace(
                        ",",
                        ""
                    )


            elif "," in value:

                parts = value.split(",")

                if len(parts[-1]) in (
                    1,
                    2
                ):

                    value = (
                        value
                        .replace(
                            ".",
                            ""
                        )
                        .replace(
                            ",",
                            "."
                        )
                    )

                else:

                    value = value.replace(
                        ",",
                        ""
                    )


            return float(value)


        except Exception:

            return None


    return text.map(
        parse_one
    )


# =========================================================
# ANÁLISIS DE DATAFRAME
# =========================================================

def dataframe_facts(
    df,
    filename,
    reference,
    facts
):

    add_fact(

        facts,

        "Cantidad de registros identificados",

        f"{len(df):,}".replace(
            ",",
            "."
        ),

        filename,

        reference,

        "structure"
    )


    add_fact(

        facts,

        "Columnas identificadas",

        ", ".join(
            str(column)
            for column in df.columns
        ),

        filename,

        reference,

        "structure"
    )


    duplicates = int(
        df.duplicated().sum()
    )


    if duplicates:

        add_fact(

            facts,

            "Filas completamente duplicadas",

            str(duplicates),

            filename,

            reference,

            "quality"
        )


    missing = int(
        df.isna()
        .sum()
        .sum()
    )


    if missing:

        add_fact(

            facts,

            "Celdas vacías identificadas",

            str(missing),

            filename,

            reference,

            "quality"
        )


    amount_added = 0


    for column in df.columns:

        column_name = str(
            column
        ).strip()


        low = column_name.lower()


        # -------------------------------------------------
        # POSIBLES COLUMNAS MONETARIAS
        # -------------------------------------------------

        if (
            amount_added < 8
            and
            any(
                keyword in low
                for keyword
                in MONEY_KEYWORDS
            )
        ):

            numbers = numeric_series(
                df[column]
            )


            valid = int(
                numbers
                .notna()
                .sum()
            )


            if (
                len(df)
                and
                valid / max(
                    len(df),
                    1
                ) >= 0.60
                and
                valid
            ):

                total = float(
                    numbers
                    .fillna(0)
                    .sum()
                )


                add_fact(

                    facts,

                    f"Total de la columna '{column_name}'",

                    f"{total:,.2f}",

                    filename,

                    reference,

                    "numeric"
                )


                amount_added += 1


        # -------------------------------------------------
        # POSIBLES COLUMNAS DE FECHA
        # -------------------------------------------------

        if any(
            keyword in low
            for keyword
            in DATE_KEYWORDS
        ):

            dates = pd.to_datetime(

                df[column],

                errors="coerce",

                dayfirst=True
            )


            valid_dates = int(
                dates
                .notna()
                .sum()
            )


            if (
                len(df)
                and
                valid_dates / max(
                    len(df),
                    1
                ) >= 0.60
                and
                valid_dates
            ):

                add_fact(

                    facts,

                    f"Rango de fechas de la columna '{column_name}'",

                    (
                        f"{dates.min().date().isoformat()}"
                        " a "
                        f"{dates.max().date().isoformat()}"
                    ),

                    filename,

                    reference,

                    "date"
                )


# =========================================================
# EXCEL / CSV
# =========================================================

def spreadsheet_facts(
    raw,
    filename,
    extension,
    facts
):

    if extension == "csv":

        df = read_csv_bytes(
            raw
        )


        dataframe_facts(

            df,

            filename,

            filename,

            facts
        )


        return


    engine = (
        "xlrd"
        if extension == "xls"
        else "openpyxl"
    )


    excel = pd.ExcelFile(

        BytesIO(raw),

        engine=engine
    )


    add_fact(

        facts,

        "Hojas identificadas en el archivo",

        ", ".join(
            excel.sheet_names
        ),

        filename,

        filename,

        "structure"
    )


    for sheet in excel.sheet_names[:12]:

        df = pd.read_excel(

            excel,

            sheet_name=sheet
        )


        dataframe_facts(

            df,

            filename,

            f"{filename} | Hoja: {sheet}",

            facts
        )


# =========================================================
# BLOQUES DE TEXTO
# =========================================================

def text_blocks(
    text,
    size=1600,
    limit=5
):

    text = clean_text(
        text
    )


    if not text:

        return []


    blocks = []

    position = 0


    while (
        position < len(text)
        and
        len(blocks) < limit
    ):

        end = min(
            position + size,
            len(text)
        )


        if end < len(text):

            cut = text.rfind(

                ". ",

                position,

                end
            )


            if cut > position + 300:

                end = cut + 1


        blocks.append(

            text[
                position:end
            ].strip()
        )


        position = end


    return [
        block
        for block in blocks
        if block
    ]


# =========================================================
# WORD
# =========================================================

def docx_facts(
    raw,
    filename,
    facts
):

    document = Document(
        BytesIO(raw)
    )


    paragraphs = [

        paragraph.text.strip()

        for paragraph
        in document.paragraphs

        if paragraph.text.strip()
    ]


    add_fact(

        facts,

        "Párrafos con contenido identificados",

        str(
            len(paragraphs)
        ),

        filename,

        filename,

        "structure"
    )


    if document.tables:

        add_fact(

            facts,

            "Tablas identificadas",

            str(
                len(document.tables)
            ),

            filename,

            filename,

            "structure"
        )


    text = "\n".join(
        paragraphs
    )


    for index, block in enumerate(
        text_blocks(text),
        start=1
    ):

        add_fact(

            facts,

            f"Extracto textual identificado {index}",

            block,

            filename,

            f"{filename} | Texto extraído",

            "text"
        )


# =========================================================
# PDF
# =========================================================

def pdf_facts(
    raw,
    filename,
    facts
):

    reader = PdfReader(
        BytesIO(raw)
    )


    add_fact(

        facts,

        "Cantidad de páginas identificadas",

        str(
            len(reader.pages)
        ),

        filename,

        filename,

        "structure"
    )


    extracted = []


    for page_number, page in enumerate(

        reader.pages[:20],

        start=1
    ):

        try:

            text = clean_text(

                page.extract_text()
                or ""
            )

        except Exception:

            text = ""


        if text:

            extracted.append(

                (
                    page_number,
                    text
                )
            )


    if not extracted:

        add_fact(

            facts,

            "Resultado de extracción de texto",

            (
                "No se pudo extraer texto del PDF. "
                "Puede tratarse de un documento escaneado."
            ),

            filename,

            filename,

            "warning"
        )


        return


    for page_number, text in extracted[:5]:

        add_fact(

            facts,

            (
                "Extracto textual identificado "
                f"- página {page_number}"
            ),

            text[:1800],

            filename,

            (
                f"{filename} | "
                f"Página {page_number}"
            ),

            "text"
        )


# =========================================================
# TXT
# =========================================================

def txt_facts(
    raw,
    filename,
    facts
):

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

        raise ValueError(
            "No se pudo decodificar el archivo de texto."
        )


    for index, block in enumerate(
        text_blocks(text),
        start=1
    ):

        add_fact(

            facts,

            f"Extracto textual identificado {index}",

            block,

            filename,

            filename,

            "text"
        )


# =========================================================
# HOME
# =========================================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


# =========================================================
# ANALIZAR DOCUMENTACIÓN
# =========================================================

@app.route(
    "/analyze",
    methods=["POST"]
)
def analyze_documents():

    files = request.files.getlist(
        "files"
    )


    free_text = request.form.get(
        "freeText",
        ""
    ).strip()


    if (
        not files
        and
        not free_text
    ):

        return jsonify(

            error=(
                "Cargue al menos un archivo "
                "o texto libre."
            )

        ), 400


    facts = []

    errors = []


    for uploaded in files:

        filename = (
            uploaded.filename
            or
            "archivo_sin_nombre"
        )


        extension = (

            filename
            .rsplit(
                ".",
                1
            )[-1]
            .lower()

            if "." in filename

            else ""
        )


        if (
            extension
            not in
            ALLOWED_EXTENSIONS
        ):

            errors.append(

                f"{filename}: formato no admitido."
            )


            continue


        try:

            raw = uploaded.read()


            if extension in {
                "xlsx",
                "xls",
                "csv"
            }:

                spreadsheet_facts(

                    raw,

                    filename,

                    extension,

                    facts
                )


            elif extension == "docx":

                docx_facts(

                    raw,

                    filename,

                    facts
                )


            elif extension == "pdf":

                pdf_facts(

                    raw,

                    filename,

                    facts
                )


            elif extension == "txt":

                txt_facts(

                    raw,

                    filename,

                    facts
                )


        except Exception as exc:

            errors.append(

                f"{filename}: {str(exc)}"
            )


    # -----------------------------------------------------
    # TEXTO LIBRE
    # -----------------------------------------------------

    if free_text:

        for index, block in enumerate(

            text_blocks(
                free_text
            ),

            start=1
        ):

            add_fact(

                facts,

                (
                    "Texto libre aportado "
                    f"por el auditor {index}"
                ),

                block,

                "Texto libre",

                "Ingreso manual",

                "text"
            )


    if facts:

        message = (

            f"Se extrajeron {len(facts)} "
            "elementos objetivos de la documentación. "
            "Revise cada uno antes de aceptarlo."
        )

    else:

        message = (

            "No se identificaron hechos verificables "
            "en la documentación cargada."
        )


    return jsonify({

        "facts":
            facts,

        "message":
            message,

        "errors":
            errors
    })


# =========================================================
# GENERAR MEMO
# =========================================================

@app.route(
    "/generate-memo",
    methods=["POST"]
)
def generate_memo():

    data = (
        request.get_json(
            silent=True
        )
        or {}
    )


    audit_data = data.get(
        "auditData",
        {}
    )


    validated_facts = data.get(
        "validatedFacts",
        []
    )


    tasks = data.get(
        "tasks",
        []
    )


    results = data.get(
        "results",
        []
    )


    findings = data.get(
        "findings",
        []
    )


    style = data.get(
        "style",
        "ejecutivo"
    )


    if (
        not validated_facts
    ):

        return jsonify(

            error=(
                "Debe existir al menos un hecho "
                "validado por el auditor."
            )

        ), 400


    memo = {

        "header": {

            "titulo":
                audit_data.get(
                    "titulo",
                    "Auditoría"
                ),

            "analisis":
                audit_data.get(
                    "analisis",
                    ""
                ),

            "sector":
                audit_data.get(
                    "sector",
                    ""
                ),

            "proceso":
                audit_data.get(
                    "proceso",
                    ""
                ),

            "periodo":
                audit_data.get(
                    "periodo",
                    ""
                ),

            "alcance":
                audit_data.get(
                    "alcance",
                    ""
                ),

            "auditor":
                audit_data.get(
                    "auditor",
                    ""
                ),

            "fecha":
                audit_data.get(

                    "fecha",

                    datetime
                        .now()
                        .strftime(
                            "%d/%m/%Y"
                        )
                )
        },


        "objetivos":
            audit_data.get(
                "objetivos",
                []
            ),


        "fuentes":
            audit_data.get(
                "fuentes",
                []
            ),


        "hechos_validados":
            validated_facts,


        "tareas":
            tasks,


        "resultados":
            results,


        "hallazgos":
            findings,


        "estilo":
            style,


        "conclusiones": (

            f"Se registraron {len(findings)} "
            "hallazgo(s) en el trabajo. "

            "La conclusión final debe ser revisada "
            "y completada por el auditor sobre la base "
            "de los hechos aceptados y la evidencia disponible."
        )
    }


    return jsonify({

        "memo":
            memo,

        "message":
            (
                "Borrador de memo generado "
                "con la información validada."
            )
    })


# =========================================================
# EXCEL - AUXILIAR
# =========================================================

def autosize_sheet(
    worksheet,
    max_width=55
):

    for column_cells in worksheet.columns:

        length = 0


        letter = (
            column_cells[0]
            .column_letter
        )


        for cell in column_cells:

            value = (

                ""
                if cell.value is None

                else str(
                    cell.value
                )
            )


            length = max(

                length,

                min(
                    len(value),
                    max_width
                )
            )


            cell.alignment = Alignment(

                vertical="top",

                wrap_text=True
            )


        worksheet
            .column_dimensions[
                letter
            ]
            .width = max(

                12,

                min(
                    length + 2,
                    max_width
                )
            )


# =========================================================
# EXPORTAR EXCEL
# =========================================================

@app.route(
    "/export-excel",
    methods=["POST"]
)
def export_excel():

    data = (
        request.get_json(
            silent=True
        )
        or {}
    )


    memo_data = data.get(
        "memo",
        {}
    )


    output = BytesIO()


    blue_dark =
        "17365D"


    blue_light =
        "EAF2F8"


    white =
        "FFFFFF"


    thin = Side(

        style="thin",

        color="D9E1E8"
    )


    with pd.ExcelWriter(

        output,

        engine="openpyxl"

    ) as writer:


        header = memo_data.get(
            "header",
            {}
        )


        # -------------------------------------------------
        # MEMO
        # -------------------------------------------------

        memo_rows = [

            [
                "TÍTULO",
                header.get(
                    "titulo",
                    ""
                )
            ],

            [
                "ANÁLISIS",
                header.get(
                    "analisis",
                    ""
                )
            ],

            [
                "SECTOR",
                header.get(
                    "sector",
                    ""
                )
            ],

            [
                "PROCESO",
                header.get(
                    "proceso",
                    ""
                )
            ],

            [
                "PERÍODO",
                header.get(
                    "periodo",
                    ""
                )
            ],

            [
                "ALCANCE",
                header.get(
                    "alcance",
                    ""
                )
            ],

            [
                "AUDITOR",
                header.get(
                    "auditor",
                    ""
                )
            ],

            [
                "FECHA",
                header.get(
                    "fecha",
                    ""
                )
            ],

            [
                "OBJETIVOS",

                "\n".join(

                    f"{index + 1}. {objective}"

                    for index, objective
                    in enumerate(
                        memo_data.get(
                            "objetivos",
                            []
                        )
                    )
                )
            ],

            [
                "CONCLUSIONES",

                memo_data.get(
                    "conclusiones",
                    ""
                )
            ]
        ]


        pd.DataFrame(

            memo_rows,

            columns=[
                "Campo",
                "Valor"
            ]

        ).to_excel(

            writer,

            sheet_name="MEMO",

            index=False
        )


        # -------------------------------------------------
        # HALLAZGOS
        # -------------------------------------------------

        findings = memo_data.get(
            "hallazgos",
            []
        )


        finding_fields = [

            "titulo",

            "descripcion",

            "condicion",

            "area_responsable",

            "responsable_plan",

            "criticidad",

            "estado",

            "fecha_objetivo",

            "riesgo",

            "propuesta_mejora",

            "fundamento_cuantitativo",

            "fuente",

            "evidencia",

            "referencia",

            "seguimiento"
        ]


        if findings:

            pd.DataFrame(

                [

                    {
                        field:
                            finding.get(
                                field,
                                ""
                            )

                        for field
                        in finding_fields
                    }

                    for finding
                    in findings
                ]

            ).to_excel(

                writer,

                sheet_name="Hallazgos",

                index=False
            )


        else:

            pd.DataFrame({

                "Mensaje": [
                    "No hay hallazgos"
                ]

            }).to_excel(

                writer,

                sheet_name="Hallazgos",

                index=False
            )


        # -------------------------------------------------
        # RESULTADOS
        # -------------------------------------------------

        results = memo_data.get(
            "resultados",
            []
        )


        if results:

            pd.DataFrame(
                results
            ).drop(

                columns=["id"],

                errors="ignore"

            ).to_excel(

                writer,

                sheet_name="Resultados",

                index=False
            )


        else:

            pd.DataFrame({

                "Mensaje": [
                    "No hay resultados"
                ]

            }).to_excel(

                writer,

                sheet_name="Resultados",

                index=False
            )


        # -------------------------------------------------
        # FUENTES
        # -------------------------------------------------

        sources = memo_data.get(
            "fuentes",
            []
        )


        if sources:

            source_rows = []


            for source in sources:

                if isinstance(
                    source,
                    dict
                ):

                    source_rows.append({

                        "Nombre":
                            source.get(
                                "name",
                                ""
                            ),

                        "Tipo":
                            source.get(
                                "type",
                                ""
                            ),

                        "Referencia":
                            source.get(
                                "reference",
                                ""
                            ),

                        "Descripción":
                            source.get(
                                "description",
                                ""
                            )
                    })


                else:

                    source_rows.append({

                        "Nombre":
                            str(source),

                        "Tipo":
                            "",

                        "Referencia":
                            "",

                        "Descripción":
                            ""
                    })


            pd.DataFrame(
                source_rows
            ).to_excel(

                writer,

                sheet_name="Fuentes",

                index=False
            )


        else:

            pd.DataFrame({

                "Mensaje": [
                    "No hay fuentes"
                ]

            }).to_excel(

                writer,

                sheet_name="Fuentes",

                index=False
            )


        # -------------------------------------------------
        # TRAZABILIDAD
        # -------------------------------------------------

        facts = memo_data.get(
            "hechos_validados",
            []
        )


        if facts:

            pd.DataFrame(

                [

                    {

                        "Descripción":
                            fact.get(
                                "description",
                                ""
                            ),

                        "Valor":
                            fact.get(
                                "value",
                                ""
                            ),

                        "Fuente":
                            fact.get(
                                "source",
                                ""
                            ),

                        "Referencia":
                            fact.get(
                                "reference",
                                ""
                            ),

                        "Tipo":
                            fact.get(
                                "kind",
                                ""
                            )
                    }

                    for fact
                    in facts
                ]

            ).to_excel(

                writer,

                sheet_name="Trazabilidad",

                index=False
            )


        else:

            pd.DataFrame({

                "Mensaje": [
                    "No hay hechos"
                ]

            }).to_excel(

                writer,

                sheet_name="Trazabilidad",

                index=False
            )


        # -------------------------------------------------
        # FORMATO
        # -------------------------------------------------

        workbook =
            writer.book


        for worksheet in workbook.worksheets:

            worksheet.freeze_panes =
                "A2"


            worksheet.auto_filter.ref =
                worksheet.dimensions


            for cell in worksheet[1]:

                cell.fill = PatternFill(

                    "solid",

                    fgColor=blue_dark
                )


                cell.font = Font(

                    color=white,

                    bold=True
                )


                cell.alignment = Alignment(

                    horizontal="center",

                    vertical="center"
                )


                cell.border = Border(

                    bottom=thin
                )


            autosize_sheet(
                worksheet
            )


        memo_sheet =
            workbook["MEMO"]


        for row in memo_sheet.iter_rows(
            min_row=2
        ):

            row[0].fill = PatternFill(

                "solid",

                fgColor=blue_light
            )


            row[0].font = Font(

                color=blue_dark,

                bold=True
            )


            for cell in row:

                cell.border = Border(

                    bottom=thin
                )


    output.seek(0)


    filename = (

        "audit_memo_"

        +

        datetime
            .now()
            .strftime(
                "%Y%m%d_%H%M%S"
            )

        +

        ".xlsx"
    )


    return send_file(

        output,

        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        as_attachment=True,

        download_name=filename
    )


# =========================================================
# MEJORAR REDACCIÓN
# =========================================================

@app.route(
    "/improve-text",
    methods=["POST"]
)
def improve_text():

    data = (
        request.get_json(
            silent=True
        )
        or {}
    )


    text = data.get(
        "text",
        ""
    )


    return jsonify({

        "original":
            text,

        "improved":
            text,

        "message": (
            "La mejora automática de redacción todavía "
            "no tiene IA conectada. "
            "El texto no fue modificado."
        )
    })


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    import os

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
