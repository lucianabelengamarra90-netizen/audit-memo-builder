import sqlite3
import os
import re
import uuid
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "audit_knowledge.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_knowledge_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS knowledge_findings (
            id TEXT PRIMARY KEY,
            category TEXT NOT NULL,
            title TEXT NOT NULL,
            situation TEXT NOT NULL,
            risk TEXT,
            proposal TEXT,
            keywords TEXT,
            source_audit TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()

    cursor.execute("SELECT COUNT(*) FROM knowledge_findings")
    count = cursor.fetchone()[0]
    if count == 0:
        seed_findings = [
            (
                str(uuid.uuid4()),
                "Inconsistencia",
                "Diferencia entre cuotas registradas y convenio contractual",
                "Se identificó una discrepancia en el número de cuotas informadas en la base de datos respecto a las pactadas en el convenio formal de la operación.",
                "Riesgo de facturación indebida, reclamos de comisionistas/clientes o descalce en las proyecciones de flujo de fondos.",
                "Conciliar las cuotas registradas en el sistema con el convenio legal y realizar los ajustes correspondientes en la base de datos.",
                "cuotas cuota convenio contrato esquema pagos",
                "Semilla Inicial de Auditoría"
            ),
            (
                str(uuid.uuid4()),
                "Falta de documentación",
                "Ausencia de especificación de la Tasa Nominal Anual (TNA)",
                "Se observó que los comprobantes u operaciones relevadas no explicitan la Tasa Nominal Anual (TNA) aplicada al recálculo de la deuda.",
                "Riesgo de falta de certeza jurídica sobre los intereses devengados y eventual descalce en la liquidación patrimonial.",
                "Explicitar formalmente la TNA en la documentación respaldatoria y recalcular los saldos según las normas vigentes.",
                "tna interes tasa tasa nominal anual interes devengado",
                "Semilla Inicial de Auditoría"
            ),
            (
                str(uuid.uuid4()),
                "Diferencia",
                "Diferencia no conciliada entre saldo informado y saldo recalculado",
                "Se detectó una diferencia no justificada entre la deuda informada en los registros contables/operativos y el saldo verificado por recálculo.",
                "Riesgo de distorsión en los saldos pasivos/activos y presentación inexacta en estados contables.",
                "Analizar las partidas conciliatorias, ajustar el saldo registrado al recálculo contractual e identificar la causa raíz del descalce.",
                "diferencia saldo recalculado recálculo deuda informada descalce",
                "Semilla Inicial de Auditoría"
            ),
            (
                str(uuid.uuid4()),
                "Falta de documentación",
                "Ausencia de comprobantes respaldatorios y firmas de autorización",
                "Se constató la falta de adjunción de legajos respaldatorios completos y la ausencia de firmas de los responsables correspondientes.",
                "Riesgo de invalidez legal de los acuerdos y vulnerabilidad ante inspecciones o auditorías externas.",
                "Requerir la regularización documental de los legajos y asegurar las firmas de autorización previa al desembolso/registro.",
                "falta firma respaldo comprobante legajo autorizacion",
                "Semilla Inicial de Auditoría"
            )
        ]
        cursor.executemany("""
            INSERT INTO knowledge_findings (id, category, title, situation, risk, proposal, keywords, source_audit)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, seed_findings)
        conn.commit()
    conn.close()


def extract_keywords(text):
    if not text:
        return ""
    words = re.findall(r"\b[a-zA-ZáéíóúüñÁÉÍÓÚÜÑ]{4,}\b", text.lower())
    ignored = {"para", "como", "esta", "este", "entre", "sobre", "desde", "hasta", "donde", "cuando", "pero", "sino"}
    filtered = [w for w in words if w not in ignored]
    return " ".join(sorted(set(filtered)))


def save_learned_findings(findings, source_audit=""):
    if not findings or not isinstance(findings, list):
        return 0
    
    conn = get_db()
    cursor = conn.cursor()
    saved = 0

    for f in findings:
        title = (f.get("title") or "").strip()
        situation = (f.get("situation") or f.get("text") or "").strip()
        category = (f.get("category") or "Hallazgo").strip()
        risk = (f.get("risk") or "").strip()
        proposal = (f.get("proposal") or "").strip()

        if not title or not situation or len(situation) < 15:
            continue

        cursor.execute("SELECT id FROM knowledge_findings WHERE LOWER(title) = ? OR LOWER(situation) = ?",
                       (title.lower(), situation.lower()))
        if cursor.fetchone():
            continue

        keywords = extract_keywords(f"{title} {situation} {category} {risk}")
        item_id = str(uuid.uuid4())

        cursor.execute("""
            INSERT INTO knowledge_findings (id, category, title, situation, risk, proposal, keywords, source_audit)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (item_id, category, title, situation, risk, proposal, keywords, source_audit or "Auditoría"))
        saved += 1

    conn.commit()
    conn.close()
    return saved


def get_relevant_knowledge(text, category="", limit=3):
    init_knowledge_db()
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM knowledge_findings")
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return []

    input_kw = set(extract_keywords(f"{text} {category}").split())
    scored = []

    for r in rows:
        item = dict(r)
        item_kw = set((item.get("keywords") or "").split())
        score = len(input_kw.intersection(item_kw))
        if category and item.get("category", "").lower() == category.lower():
            score += 2
        scored.append((score, item))

    scored.sort(key=lambda x: x[0], reverse=True)
    return [item for score, item in scored[:limit]]


def get_all_knowledge():
    init_knowledge_db()
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM knowledge_findings ORDER BY created_at DESC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def add_custom_knowledge_item(category, title, situation, risk="", proposal="", source_audit="Regla Manual"):
    init_knowledge_db()
    conn = get_db()
    cursor = conn.cursor()
    item_id = str(uuid.uuid4())
    keywords = extract_keywords(f"{title} {situation} {category} {risk}")
    cursor.execute("""
        INSERT INTO knowledge_findings (id, category, title, situation, risk, proposal, keywords, source_audit)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (item_id, category, title, situation, risk, proposal, keywords, source_audit))
    conn.commit()
    conn.close()
    return item_id


def delete_knowledge_item(item_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM knowledge_findings WHERE id = ?", (item_id,))
    deleted = cursor.rowcount > 0
    conn.commit()
    conn.close()
    return deleted
