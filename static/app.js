// ============================================================
// AUDIT MEMO BUILDER
// static/app.js
//
// Objetivo:
// - Mantener cada auditoría aislada.
// - Evitar contaminación entre papeles de trabajo.
// - Conservar trazabilidad.
// - Extraer información sin convertirla automáticamente
//   en hallazgo.
// ============================================================


// ============================================================
// CONFIGURACIÓN
// ============================================================

const STORAGE_PREFIX = "auditMemoBuilder";

const AUTOSAVE_DELAY = 400;

let autosaveTimer = null;


// ============================================================
// ESTADO
// ============================================================

function createEmptyState() {

    return {

        auditId: crypto.randomUUID(),

        createdAt: new Date().toISOString(),

        updatedAt: new Date().toISOString(),

        currentStep: 1,

        general: {
            title: "",
            area: "",
            process: "",
            period: "",
            auditor: "",
            objective: "",
            scope: ""
        },

        files: [],

        sources: [],

        extracted: [],

        findings: []
    };
}


let state = createEmptyState();


// ============================================================
// STORAGE
// ============================================================

function getStorageKey(auditId = state.auditId) {

    return `${STORAGE_PREFIX}:${auditId}`;
}


function getCurrentAuditKey() {

    return `${STORAGE_PREFIX}:currentAudit`;
}


function saveState() {

    try {

        state.updatedAt = new Date().toISOString();

        localStorage.setItem(
            getStorageKey(),
            JSON.stringify(state)
        );

        localStorage.setItem(
            getCurrentAuditKey(),
            state.auditId
        );

    } catch (error) {

        console.warn(
            "No se pudo guardar el estado:",
            error
        );
    }
}


function scheduleSave() {

    clearTimeout(
        autosaveTimer
    );

    autosaveTimer = setTimeout(
        saveState,
        AUTOSAVE_DELAY
    );
}


function loadCurrentAudit() {

    try {

        const currentAuditId = localStorage.getItem(
            getCurrentAuditKey()
        );

        if (!currentAuditId) {
            return false;
        }

        const saved = localStorage.getItem(
            getStorageKey(currentAuditId)
        );

        if (!saved) {
            return false;
        }

        const parsed = JSON.parse(
            saved
        );

        if (
            !parsed ||
            typeof parsed !== "object"
        ) {
            return false;
        }

        state = {
            ...createEmptyState(),
            ...parsed,
            general: {
                ...createEmptyState().general,
                ...(parsed.general || {})
            },
            files: parsed.files || [],
            sources: parsed.sources || [],
            extracted: parsed.extracted || [],
            findings: parsed.findings || []
        };

        return true;

    } catch (error) {

        console.warn(
            "No se pudo recuperar la auditoría:",
            error
        );

        return false;
    }
}


// ============================================================
// NUEVA AUDITORÍA
// ============================================================

function hasAuditContent() {

    const hasGeneral = Object.values(
        state.general
    ).some(
        value => String(value || "").trim()
    );

    return (
        hasGeneral ||
        state.files.length > 0 ||
        state.sources.length > 0 ||
        state.extracted.length > 0 ||
        state.findings.length > 0
    );
}


function startNewAudit({
    keepAuditor = true
} = {}) {

    const previousAuditor = (
        keepAuditor
        ? state.general.auditor
        : ""
    );

    state = createEmptyState();

    state.general.auditor = previousAuditor;

    saveState();

    hydrateForm();

    renderFiles();

    renderExtraction();

    renderFindings();

    renderMemoPreview();

    updateFinalSummary();

    goToStep(1);

    showToast(
        "Se inició una auditoría nueva.",
        "success"
    );
}


function requestNewAudit() {

    if (!hasAuditContent()) {

        startNewAudit();

        return;
    }

    const confirmed = window.confirm(
        "¿Querés iniciar una auditoría nueva?\n\n" +
        "La auditoría actual quedará guardada en este navegador, " +
        "pero la nueva comenzará sin hallazgos, fuentes ni extracciones anteriores."
    );

    if (!confirmed) {
        return;
    }

    startNewAudit();
}


// ============================================================
// ELEMENTOS GENERALES
// ============================================================

function getElement(id) {

    return document.getElementById(
        id
    );
}


function getValue(id) {

    const element = getElement(id);

    if (!element) {
        return "";
    }

    return (
        element.value ||
        ""
    ).trim();
}


function setValue(id, value) {

    const element = getElement(id);

    if (!element) {
        return;
    }

    element.value = value || "";
}


function escapeHtml(value) {

    return String(
        value ?? ""
    )
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}


// ============================================================
// FORMULARIO GENERAL
// ============================================================

const GENERAL_FIELDS = {

    auditTitle: "title",

    auditArea: "area",

    auditProcess: "process",

    auditPeriod: "period",

    auditAuditor: "auditor",

    auditObjective: "objective",

    auditScope: "scope"
};


function hydrateForm() {

    Object.entries(
        GENERAL_FIELDS
    ).forEach(
        ([elementId, stateKey]) => {

            setValue(
                elementId,
                state.general[stateKey]
            );
        }
    );
}


function bindGeneralFields() {

    Object.entries(
        GENERAL_FIELDS
    ).forEach(
        ([elementId, stateKey]) => {

            const element = getElement(
                elementId
            );

            if (!element) {
                return;
            }

            element.addEventListener(
                "input",
                () => {

                    state.general[stateKey] = (
                        element.value ||
                        ""
                    );

                    scheduleSave();

                    renderMemoPreview();

                    updateFinalSummary();
                }
            );
        }
    );
}


// ============================================================
// ARCHIVOS
// ============================================================

let selectedFiles = [];


function fileSignature(file) {

    return [
        file.name,
        file.size,
        file.lastModified
    ].join("::");
}


function addFiles(fileList) {

    const incoming = Array.from(
        fileList || []
    );

    if (!incoming.length) {
        return;
    }

    // --------------------------------------------------------
    // CONTROL DE CONTAMINACIÓN ENTRE AUDITORÍAS
    //
    // Si ya existen hallazgos o extracciones y el usuario
    // carga papeles nuevos, preguntamos si corresponden
    // a una auditoría nueva.
    // --------------------------------------------------------

    if (
        (
            state.findings.length > 0 ||
            state.extracted.length > 0
        ) &&
        selectedFiles.length === 0
    ) {

        const startClean = window.confirm(
            "La auditoría actual ya contiene extracciones o hallazgos.\n\n" +
            "¿Estos archivos corresponden a una auditoría NUEVA?\n\n" +
            "Aceptar = comenzar un memo limpio.\n" +
            "Cancelar = agregar los archivos a la auditoría actual."
        );

        if (startClean) {

            startNewAudit({
                keepAuditor: true
            });
        }
    }

    const currentSignatures = new Set(
        selectedFiles.map(
            fileSignature
        )
    );

    incoming.forEach(
        file => {

            const signature = fileSignature(
                file
            );

            if (
                !currentSignatures.has(
                    signature
                )
            ) {

                selectedFiles.push(
                    file
                );

                currentSignatures.add(
                    signature
                );
            }
        }
    );

    syncFileMetadata();

    renderFiles();

    scheduleSave();
}


function syncFileMetadata() {

    state.files = selectedFiles.map(
        file => ({
            name: file.name,
            size: file.size,
            type: file.type || "",
            lastModified: file.lastModified
        })
    );

    state.sources = state.files.map(
        file => ({
            name: file.name,
            type: detectSourceType(
                file.name
            ),
            reference: "",
            description: "Papel de trabajo aportado para la auditoría."
        })
    );
}


function detectSourceType(filename) {

    const extension = (
        filename.split(".").pop() ||
        ""
    ).toLowerCase();

    const types = {
        xlsx: "Excel",
        csv: "CSV",
        docx: "Word",
        pdf: "PDF",
        txt: "TXT"
    };

    return types[extension] || extension.toUpperCase();
}


function removeFile(index) {

    selectedFiles.splice(
        index,
        1
    );

    syncFileMetadata();

    renderFiles();

    scheduleSave();
}


function renderFiles() {

    const container = (
        getElement("selectedFiles") ||
        getElement("fileList")
    );

    if (!container) {
        return;
    }

    if (!selectedFiles.length) {

        container.innerHTML = "";

        return;
    }

    container.innerHTML = selectedFiles.map(
        (file, index) => `

            <div class="file-item">

                <div class="file-info">

                    <strong>
                        ${escapeHtml(file.name)}
                    </strong>

                    <span>
                        ${formatFileSize(file.size)}
                    </span>

                </div>

                <button
                    type="button"
                    class="icon-button"
                    onclick="removeFile(${index})"
                    title="Quitar archivo"
                >
                    ×
                </button>

            </div>

        `
    ).join("");
}


function formatFileSize(bytes) {

    if (!bytes) {
        return "0 KB";
    }

    const mb = bytes / (
        1024 * 1024
    );

    if (mb >= 1) {
        return `${mb.toFixed(1)} MB`;
    }

    return `${(
        bytes / 1024
    ).toFixed(0)} KB`;
}


// ============================================================
// DRAG & DROP
// ============================================================

function setupDropzone() {

    const input = getElement(
        "fileInput"
    );

    const dropzone = getElement(
        "dropzone"
    );

    if (input) {

        input.addEventListener(
            "change",
            event => {

                addFiles(
                    event.target.files
                );

                input.value = "";
            }
        );
    }

    if (!dropzone) {
        return;
    }

    ["dragenter", "dragover"].forEach(
        eventName => {

            dropzone.addEventListener(
                eventName,
                event => {

                    event.preventDefault();

                    dropzone.classList.add(
                        "dragging"
                    );
                }
            );
        }
    );

    ["dragleave", "drop"].forEach(
        eventName => {

            dropzone.addEventListener(
                eventName,
                event => {

                    event.preventDefault();

                    dropzone.classList.remove(
                        "dragging"
                    );
                }
            );
        }
    );

    dropzone.addEventListener(
        "drop",
        event => {

            addFiles(
                event.dataTransfer.files
            );
        }
    );
}


// ============================================================
// EXTRACCIÓN
// ============================================================

async function extractInformation() {

    const freeText = getValue(
        "freeText"
    );

    if (
        selectedFiles.length === 0 &&
        !freeText
    ) {

        showToast(
            "Cargá al menos un papel de trabajo o ingresá texto.",
            "warning"
        );

        return;
    }

    const button = (
        getElement("extractButton") ||
        getElement("analyzeButton")
    );

    setButtonLoading(
        button,
        true,
        "Extrayendo..."
    );

    const formData = new FormData();

    selectedFiles.forEach(
        file => {

            formData.append(
                "files",
                file
            );
        }
    );

    formData.append(
        "freeText",
        freeText
    );

    try {

        const response = await fetch(
            "/extract",
            {
                method: "POST",
                body: formData
            }
        );

        let data = {};

        try {

            data = await response.json();

        } catch {

            data = {};
        }

        if (!response.ok) {

            throw new Error(
                data.error ||
                "No se pudo procesar la documentación."
            );
        }

        // ----------------------------------------------------
        // IMPORTANTE:
        // reemplazamos la extracción de ESTA auditoría.
        // No concatenamos basura de auditorías anteriores.
        // ----------------------------------------------------

        state.extracted = (
            data.items ||
            []
        ).map(
            item => ({
                ...item,
                included: Boolean(
                    item.included
                ),
                converted: Boolean(
                    item.converted
                )
            })
        );

        syncFileMetadata();

        saveState();

        renderExtraction();

        updateFinalSummary();

        goToStep(2);

        const errors = (
            data.errors ||
            []
        );

        const warnings = (
            data.warnings ||
            []
        );

        let message = (
            data.message ||
            `Se detectaron ${state.extracted.length} elementos.`
        );

        if (warnings.length) {

            message += (
                ` Advertencias: ${warnings.length}.`
            );
        }

        if (errors.length) {

            message += (
                ` Archivos con error: ${errors.length}.`
            );
        }

        showToast(
            message,
            errors.length
                ? "warning"
                : "success"
        );

    } catch (error) {

        console.error(
            error
        );

        showToast(
            error.message ||
            "Se produjo un error durante la extracción.",
            "error"
        );

    } finally {

        setButtonLoading(
            button,
            false
        );
    }
}


// ============================================================
// RENDER EXTRACCIÓN
// ============================================================

function renderExtraction() {

    const container = (
        getElement("extractionResults") ||
        getElement("extractedItems")
    );

    if (!container) {
        return;
    }

    if (!state.extracted.length) {

        container.innerHTML = `

            <div class="empty-state">

                <h3>
                    Sin elementos extraídos
                </h3>

                <p>
                    Cargá los papeles de trabajo y ejecutá la extracción.
                </p>

            </div>

        `;

        return;
    }

    container.innerHTML = state.extracted.map(
        (item, index) => {

            const checked = item.included
                ? "checked"
                : "";

            const converted = item.converted
                ? "Convertido en hallazgo"
                : "Disponible";

            return `

                <article class="extraction-card">

                    <div class="extraction-card-header">

                        <div>

                            <span class="category-badge">
                                ${escapeHtml(item.category)}
                            </span>

                            <strong class="source-title">
                                ${escapeHtml(item.filename)}
                            </strong>

                        </div>

                        <label class="include-check">

                            <input
                                type="checkbox"
                                ${checked}
                                onchange="toggleExtractionIncluded(
                                    ${index},
                                    this.checked
                                )"
                            >

                            Incluir

                        </label>

                    </div>

                    <textarea
                        class="extraction-text"
                        oninput="updateExtractionText(
                            ${index},
                            this.value
                        )"
                    >${escapeHtml(item.text)}</textarea>

                    <div class="trace-meta">

                        ${
                            item.originName
                            ? `
                                <span>
                                    Solapa / origen:
                                    <strong>
                                        ${escapeHtml(item.originName)}
                                    </strong>
                                </span>
                            `
                            : ""
                        }

                        ${
                            item.reference
                            ? `
                                <span>
                                    ${escapeHtml(item.reference)}
                                </span>
                            `
                            : ""
                        }

                        ${
                            item.keyword
                            ? `
                                <span>
                                    Keyword:
                                    ${escapeHtml(item.keyword)}
                                </span>
                            `
                            : ""
                        }

                    </div>

                    <div class="extraction-card-footer">

                        <span class="conversion-status">
                            ${converted}
                        </span>

                        <button
                            type="button"
                            class="secondary-button"
                            onclick="convertExtractionToFinding(${index})"
                            ${item.converted ? "disabled" : ""}
                        >
                            Convertir en hallazgo
                        </button>

                    </div>

                </article>

            `;
        }
    ).join("");
}


function toggleExtractionIncluded(
    index,
    included
) {

    const item = state.extracted[
        index
    ];

    if (!item) {
        return;
    }

    item.included = included;

    scheduleSave();

    updateFinalSummary();
}


function updateExtractionText(
    index,
    value
) {

    const item = state.extracted[
        index
    ];

    if (!item) {
        return;
    }

    item.text = value;

    scheduleSave();
}


// ============================================================
// CONVERTIR A HALLAZGO
// ============================================================

function convertExtractionToFinding(index) {

    const item = state.extracted[
        index
    ];

    if (!item) {
        return;
    }

    if (item.converted) {

        showToast(
            "Este elemento ya fue convertido.",
            "warning"
        );

        return;
    }

    const finding = {

        id: crypto.randomUUID(),

        title: buildFindingTitle(
            item
        ),

        situation: (
            item.text ||
            ""
        ),

        risk: "",

        proposal: "",

        responsibleArea: "",

        actionOwner: "",

        severity: "Media",

        status: "Pendiente",

        targetDate: "",

        quantitativeBasis: "",

        sourceFile: (
            item.filename ||
            ""
        ),

        sourceLocation: (
            item.originName ||
            ""
        ),

        evidence: (
            item.reference ||
            ""
        ),

        ticket: "",

        followUp: "",

        sourceItemId: item.id
    };

    state.findings.push(
        finding
    );

    item.converted = true;

    item.included = true;

    saveState();

    renderExtraction();

    renderFindings();

    renderMemoPreview();

    updateFinalSummary();

    showToast(
        "Se creó un hallazgo para revisión.",
        "success"
    );
}


function buildFindingTitle(item) {

    const category = (
        item.category ||
        "Hallazgo"
    );

    const origin = (
        item.originName ||
        ""
    );

    if (origin) {

        return `${category} - ${origin}`;
    }

    return category;
}


// ============================================================
// HALLAZGOS
// ============================================================

function addBlankFinding() {

    state.findings.push({

        id: crypto.randomUUID(),

        title: "",

        situation: "",

        risk: "",

        proposal: "",

        responsibleArea: "",

        actionOwner: "",

        severity: "Media",

        status: "Pendiente",

        targetDate: "",

        quantitativeBasis: "",

        sourceFile: "",

        sourceLocation: "",

        evidence: "",

        ticket: "",

        followUp: "",

        sourceItemId: null
    });

    saveState();

    renderFindings();

    goToStep(3);
}


function updateFinding(
    index,
    field,
    value
) {

    const finding = state.findings[
        index
    ];

    if (!finding) {
        return;
    }

    finding[field] = value;

    scheduleSave();

    renderMemoPreview();

    updateFinalSummary();
}


function deleteFinding(index) {

    const finding = state.findings[
        index
    ];

    if (!finding) {
        return;
    }

    const confirmed = window.confirm(
        `¿Eliminar el Hallazgo ${String(
            index + 1
        ).padStart(2, "0")}?`
    );

    if (!confirmed) {
        return;
    }

    if (finding.sourceItemId) {

        const sourceItem = state.extracted.find(
            item => (
                item.id === finding.sourceItemId
            )
        );

        if (sourceItem) {

            sourceItem.converted = false;
        }
    }

    state.findings.splice(
        index,
        1
    );

    saveState();

    renderExtraction();

    renderFindings();

    renderMemoPreview();

    updateFinalSummary();
}


function moveFinding(
    index,
    direction
) {

    const targetIndex = (
        index + direction
    );

    if (
        targetIndex < 0 ||
        targetIndex >= state.findings.length
    ) {
        return;
    }

    const temporary = state.findings[
        index
    ];

    state.findings[index] = state.findings[
        targetIndex
    ];

    state.findings[targetIndex] = temporary;

    saveState();

    renderFindings();

    renderMemoPreview();
}


// ============================================================
// RENDER HALLAZGOS
// ============================================================

function renderFindings() {

    const container = getElement(
        "findingsContainer"
    );

    if (!container) {
        return;
    }

    if (!state.findings.length) {

        container.innerHTML = `

            <div class="empty-state">

                <h3>
                    Todavía no hay hallazgos
                </h3>

                <p>
                    Convertí elementos desde Extracción
                    o agregá un hallazgo manual.
                </p>

            </div>

        `;

        return;
    }

    container.innerHTML = state.findings.map(
        (finding, index) => {

            const number = String(
                index + 1
            ).padStart(
                2,
                "0"
            );

            const severityClass = (
                finding.severity ||
                "Media"
            ).toLowerCase();

            return `

                <article
                    class="finding-card severity-${escapeHtml(severityClass)}"
                >

                    <div class="finding-header">

                        <div>

                            <span class="finding-number">
                                Hallazgo ${number}
                            </span>

                            <span class="severity-pill severity-${escapeHtml(severityClass)}">
                                ${escapeHtml(finding.severity || "Media")}
                            </span>

                        </div>

                        <div class="finding-actions">

                            <button
                                type="button"
                                class="icon-button"
                                onclick="moveFinding(${index}, -1)"
                                title="Subir"
                            >
                                ↑
                            </button>

                            <button
                                type="button"
                                class="icon-button"
                                onclick="moveFinding(${index}, 1)"
                                title="Bajar"
                            >
                                ↓
                            </button>

                            <button
                                type="button"
                                class="icon-button danger"
                                onclick="deleteFinding(${index})"
                                title="Eliminar"
                            >
                                ×
                            </button>

                        </div>

                    </div>

                    <div class="form-field">

                        <label>
                            Título
                        </label>

                        <input
                            type="text"
                            value="${escapeHtml(finding.title)}"
                            oninput="updateFinding(
                                ${index},
                                'title',
                                this.value
                            )"
                        >

                    </div>

                    <div class="form-field">

                        <label>
                            Situación observada
                        </label>

                        <textarea
                            oninput="updateFinding(
                                ${index},
                                'situation',
                                this.value
                            )"
                        >${escapeHtml(finding.situation)}</textarea>

                    </div>

                    <div class="form-grid two-columns">

                        <div class="form-field">

                            <label>
                                Riesgo
                            </label>

                            <textarea
                                oninput="updateFinding(
                                    ${index},
                                    'risk',
                                    this.value
                                )"
                            >${escapeHtml(finding.risk)}</textarea>

                        </div>

                        <div class="form-field">

                            <label>
                                Propuesta de mejora
                            </label>

                            <textarea
                                oninput="updateFinding(
                                    ${index},
                                    'proposal',
                                    this.value
                                )"
                            >${escapeHtml(finding.proposal)}</textarea>

                        </div>

                    </div>

                    <div class="form-grid three-columns">

                        <div class="form-field">

                            <label>
                                Área responsable
                            </label>

                            <input
                                value="${escapeHtml(finding.responsibleArea)}"
                                oninput="updateFinding(
                                    ${index},
                                    'responsibleArea',
                                    this.value
                                )"
                            >

                        </div>

                        <div class="form-field">

                            <label>
                                Criticidad
                            </label>

                            <select
                                onchange="updateFinding(
                                    ${index},
                                    'severity',
                                    this.value
                                ); renderFindings();"
                            >

                                <option
                                    value="Alta"
                                    ${finding.severity === "Alta" ? "selected" : ""}
                                >
                                    Alta
                                </option>

                                <option
                                    value="Media"
                                    ${finding.severity === "Media" ? "selected" : ""}
                                >
                                    Media
                                </option>

                                <option
                                    value="Baja"
                                    ${finding.severity === "Baja" ? "selected" : ""}
                                >
                                    Baja
                                </option>

                            </select>

                        </div>

                        <div class="form-field">

                            <label>
                                Estado
                            </label>

                            <input
                                value="${escapeHtml(finding.status)}"
                                oninput="updateFinding(
                                    ${index},
                                    'status',
                                    this.value
                                )"
                            >

                        </div>

                    </div>

                    <details class="finding-details">

                        <summary>
                            Seguimiento y trazabilidad
                        </summary>

                        <div class="form-grid two-columns">

                            <div class="form-field">

                                <label>
                                    Responsable del plan
                                </label>

                                <input
                                    value="${escapeHtml(finding.actionOwner)}"
                                    oninput="updateFinding(
                                        ${index},
                                        'actionOwner',
                                        this.value
                                    )"
                                >

                            </div>

                            <div class="form-field">

                                <label>
                                    Fecha compromiso
                                </label>

                                <input
                                    type="date"
                                    value="${escapeHtml(finding.targetDate)}"
                                    oninput="updateFinding(
                                        ${index},
                                        'targetDate',
                                        this.value
                                    )"
                                >

                            </div>

                            <div class="form-field">

                                <label>
                                    Base cuantitativa
                                </label>

                                <input
                                    value="${escapeHtml(finding.quantitativeBasis)}"
                                    oninput="updateFinding(
                                        ${index},
                                        'quantitativeBasis',
                                        this.value
                                    )"
                                >

                            </div>

                            <div class="form-field">

                                <label>
                                    Ticket
                                </label>

                                <input
                                    value="${escapeHtml(finding.ticket)}"
                                    oninput="updateFinding(
                                        ${index},
                                        'ticket',
                                        this.value
                                    )"
                                >

                            </div>

                            <div class="form-field">

                                <label>
                                    Archivo de origen
                                </label>

                                <input
                                    value="${escapeHtml(finding.sourceFile)}"
                                    oninput="updateFinding(
                                        ${index},
                                        'sourceFile',
                                        this.value
                                    )"
                                >

                            </div>

                            <div class="form-field">

                                <label>
                                    Solapa / origen
                                </label>

                                <input
                                    value="${escapeHtml(finding.sourceLocation)}"
                                    oninput="updateFinding(
                                        ${index},
                                        'sourceLocation',
                                        this.value
                                    )"
                                >

                            </div>

                        </div>

                        <div class="form-field">

                            <label>
                                Evidencia / referencia
                            </label>

                            <textarea
                                oninput="updateFinding(
                                    ${index},
                                    'evidence',
                                    this.value
                                )"
                            >${escapeHtml(finding.evidence)}</textarea>

                        </div>

                        <div class="form-field">

                            <label>
                                Seguimiento
                            </label>

                            <textarea
                                oninput="updateFinding(
                                    ${index},
                                    'followUp',
                                    this.value
                                )"
                            >${escapeHtml(finding.followUp)}</textarea>

                        </div>

                    </details>

                    ${
                        finding.sourceFile
                        ? `
                            <div class="finding-source">
                                Origen:
                                <strong>
                                    ${escapeHtml(finding.sourceFile)}
                                </strong>
                                ${
                                    finding.sourceLocation
                                    ? ` · Solapa: ${escapeHtml(finding.sourceLocation)}`
                                    : ""
                                }
                            </div>
                        `
                        : ""
                    }

                </article>

            `;
        }
    ).join("");
}


// ============================================================
// CONTROL DE TRAZABILIDAD ANTES DE EXPORTAR
// ============================================================

function validateFindingSources() {

    const activeFiles = new Set(
        state.sources.map(
            source => (
                source.name ||
                ""
            ).trim()
        ).filter(Boolean)
    );

    const foreignFindings = state.findings.filter(
        finding => {

            const sourceFile = (
                finding.sourceFile ||
                ""
            ).trim();

            if (!sourceFile) {
                return false;
            }

            return !activeFiles.has(
                sourceFile
            );
        }
    );

    return foreignFindings;
}


// ============================================================
// MEMO PREVIEW
// ============================================================

function renderMemoPreview() {

    const container = getElement(
        "memoPreview"
    );

    if (!container) {
        return;
    }

    const findingsHtml = state.findings.length
        ? state.findings.map(
            (finding, index) => `

                <section class="memo-finding">

                    <h4>
                        Hallazgo ${String(index + 1).padStart(2, "0")}
                        ${finding.title ? ` - ${escapeHtml(finding.title)}` : ""}
                    </h4>

                    <p>
                        <strong>Situación observada:</strong>
                        ${escapeHtml(finding.situation)}
                    </p>

                    <p>
                        <strong>Riesgo:</strong>
                        ${escapeHtml(finding.risk)}
                    </p>

                    <p>
                        <strong>Propuesta de mejora:</strong>
                        ${escapeHtml(finding.proposal)}
                    </p>

                </section>

            `
        ).join("")
        : `
            <p class="muted">
                No se incorporaron hallazgos.
            </p>
        `;

    container.innerHTML = `

        <div class="memo-preview-document">

            <h2>
                ${escapeHtml(
                    state.general.title ||
                    "Memo de Auditoría"
                )}
            </h2>

            <div class="memo-meta">

                <p>
                    <strong>Área:</strong>
                    ${escapeHtml(state.general.area)}
                </p>

                <p>
                    <strong>Proceso:</strong>
                    ${escapeHtml(state.general.process)}
                </p>

                <p>
                    <strong>Período:</strong>
                    ${escapeHtml(state.general.period)}
                </p>

                <p>
                    <strong>Auditor:</strong>
                    ${escapeHtml(state.general.auditor)}
                </p>

            </div>

            <h3>
                Objetivo
            </h3>

            <p>
                ${escapeHtml(state.general.objective)}
            </p>

            <h3>
                Alcance
            </h3>

            <p>
                ${escapeHtml(state.general.scope)}
            </p>

            <h3>
                Hallazgos
            </h3>

            ${findingsHtml}

        </div>

    `;
}


// ============================================================
// VALIDACIÓN FINAL
// ============================================================

function validateMemo() {

    const issues = [];

    if (!state.general.title.trim()) {

        issues.push(
            "Falta indicar el nombre de la auditoría."
        );
    }

    if (!state.general.objective.trim()) {

        issues.push(
            "Falta completar el objetivo."
        );
    }

    state.findings.forEach(
        (finding, index) => {

            const number = String(
                index + 1
            ).padStart(
                2,
                "0"
            );

            if (!finding.title.trim()) {

                issues.push(
                    `Hallazgo ${number}: falta título.`
                );
            }

            if (!finding.situation.trim()) {

                issues.push(
                    `Hallazgo ${number}: falta Situación observada.`
                );
            }
        }
    );

    const foreignFindings = validateFindingSources();

    if (foreignFindings.length) {

        issues.push(
            (
                `${foreignFindings.length} hallazgo(s) contienen ` +
                "archivos de origen que no pertenecen a las fuentes " +
                "actualmente cargadas."
            )
        );
    }

    return issues;
}


// ============================================================
// RESUMEN FINAL
// ============================================================

function updateFinalSummary() {

    const container = getElement(
        "finalSummary"
    );

    if (!container) {
        return;
    }

    const issues = validateMemo();

    container.innerHTML = `

        <div class="summary-grid">

            <div class="summary-card">

                <strong>
                    ${state.sources.length}
                </strong>

                <span>
                    Fuentes
                </span>

            </div>

            <div class="summary-card">

                <strong>
                    ${state.extracted.length}
                </strong>

                <span>
                    Elementos extraídos
                </span>

            </div>

            <div class="summary-card">

                <strong>
                    ${state.findings.length}
                </strong>

                <span>
                    Hallazgos
                </span>

            </div>

            <div class="summary-card">

                <strong>
                    ${issues.length}
                </strong>

                <span>
                    Validaciones pendientes
                </span>

            </div>

        </div>

        ${
            issues.length
            ? `
                <div class="validation-box warning">

                    <strong>
                        Revisar antes de exportar:
                    </strong>

                    <ul>

                        ${issues.map(
                            issue => `
                                <li>
                                    ${escapeHtml(issue)}
                                </li>
                            `
                        ).join("")}

                    </ul>

                </div>
            `
            : `
                <div class="validation-box success">
                    El memo no presenta validaciones bloqueantes.
                </div>
            `
        }

    `;
}


// ============================================================
// EXPORTAR EXCEL
// ============================================================

async function exportExcel() {

    syncGeneralFromForm();

    const issues = validateMemo();

    const foreignFindings = validateFindingSources();

    // --------------------------------------------------------
    // CONTROL CRÍTICO:
    // NO exportamos hallazgos provenientes de otra auditoría.
    // --------------------------------------------------------

    if (foreignFindings.length) {

        const files = [
            ...new Set(
                foreignFindings.map(
                    finding => finding.sourceFile
                )
            )
        ];

        showToast(
            (
                "Exportación bloqueada. Hay hallazgos asociados a " +
                "fuentes que no pertenecen a esta auditoría: " +
                files.join(", ")
            ),
            "error"
        );

        return;
    }

    if (issues.length) {

        const continueExport = window.confirm(
            (
                "El memo tiene validaciones pendientes:\n\n" +
                issues.join("\n") +
                "\n\n¿Querés exportarlo igualmente?"
            )
        );

        if (!continueExport) {
            return;
        }
    }

    const button = getElement(
        "exportExcelButton"
    );

    setButtonLoading(
        button,
        true,
        "Generando..."
    );

    const payload = {

        memo: {

            auditId: state.auditId,

            general: state.general,

            findings: state.findings,

            sources: state.sources,

            extracted: state.extracted
        }
    };

    try {

        const response = await fetch(
            "/export-excel",
            {
                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify(
                    payload
                )
            }
        );

        if (!response.ok) {

            let message = (
                "No se pudo generar el Excel."
            );

            try {

                const errorData = await response.json();

                message = (
                    errorData.error ||
                    message
                );

            } catch {
                // dejamos mensaje genérico
            }

            throw new Error(
                message
            );
        }

        const blob = await response.blob();

        const url = URL.createObjectURL(
            blob
        );

        const link = document.createElement(
            "a"
        );

        link.href = url;

        const safeTitle = (
            state.general.title ||
            "Audit_Memo"
        )
            .replace(
                /[^a-z0-9áéíóúñü _-]/gi,
                ""
            )
            .trim()
            .replace(
                /\s+/g,
                "_"
            );

        link.download = (
            `${safeTitle || "Audit_Memo"}.xlsx`
        );

        document.body.appendChild(
            link
        );

        link.click();

        link.remove();

        URL.revokeObjectURL(
            url
        );

        showToast(
            "Excel generado correctamente.",
            "success"
        );

    } catch (error) {

        console.error(
            error
        );

        showToast(
            error.message ||
            "No se pudo exportar el memo.",
            "error"
        );

    } finally {

        setButtonLoading(
            button,
            false
        );
    }
}


// ============================================================
// IA - MEJORAR TEXTO
// ============================================================

async function improveField(
    elementId,
    fieldType
) {

    const element = getElement(
        elementId
    );

    if (!element) {
        return;
    }

    const text = (
        element.value ||
        ""
    ).trim();

    if (!text) {

        showToast(
            "Primero escribí una idea para mejorar.",
            "warning"
        );

        element.focus();

        return;
    }

    const buttons = document.querySelectorAll(
        ".ai-button"
    );

    buttons.forEach(
        button => button.disabled = true
    );

    element.disabled = true;

    try {

        const response = await fetch(
            "/improve-text",
            {
                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify({
                    text,
                    fieldType
                })
            }
        );

        let data = {};

        try {

            data = await response.json();

        } catch {

            data = {};
        }

        if (!response.ok) {

            throw new Error(
                data.error ||
                "No se pudo mejorar el texto."
            );
        }

        element.value = (
            data.improved ||
            text
        );

        element.dispatchEvent(
            new Event(
                "input",
                {
                    bubbles: true
                }
            )
        );

        showToast(
            "Redacción actualizada.",
            "success"
        );

    } catch (error) {

        showToast(
            error.message ||
            "No se pudo utilizar la función de IA.",
            "error"
        );

    } finally {

        element.disabled = false;

        buttons.forEach(
            button => button.disabled = false
        );
    }
}


// ============================================================
// WIZARD
// ============================================================

function goToStep(step) {

    const totalSteps = 5;

    const nextStep = Math.min(
        totalSteps,
        Math.max(
            1,
            Number(step) || 1
        )
    );

    state.currentStep = nextStep;

    document.querySelectorAll(
        "[data-step-panel]"
    ).forEach(
        panel => {

            const panelStep = Number(
                panel.dataset.stepPanel
            );

            panel.classList.toggle(
                "active",
                panelStep === nextStep
            );
        }
    );

    document.querySelectorAll(
        "[data-step]"
    ).forEach(
        stepElement => {

            const stepNumber = Number(
                stepElement.dataset.step
            );

            stepElement.classList.toggle(
                "active",
                stepNumber === nextStep
            );

            stepElement.classList.toggle(
                "completed",
                stepNumber < nextStep
            );
        }
    );

    if (nextStep === 3) {

        renderFindings();
    }

    if (nextStep === 4) {

        renderMemoPreview();
    }

    if (nextStep === 5) {

        updateFinalSummary();
    }

    scheduleSave();

    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
}


function nextStep() {

    goToStep(
        state.currentStep + 1
    );
}


function previousStep() {

    goToStep(
        state.currentStep - 1
    );
}


// ============================================================
// SYNC GENERAL
// ============================================================

function syncGeneralFromForm() {

    Object.entries(
        GENERAL_FIELDS
    ).forEach(
        ([elementId, stateKey]) => {

            const element = getElement(
                elementId
            );

            if (!element) {
                return;
            }

            state.general[stateKey] = (
                element.value ||
                ""
            );
        }
    );

    saveState();
}


// ============================================================
// BOTONES
// ============================================================

function setButtonLoading(
    button,
    loading,
    loadingText = "Procesando..."
) {

    if (!button) {
        return;
    }

    if (loading) {

        button.dataset.originalText = (
            button.innerHTML
        );

        button.disabled = true;

        button.innerHTML = loadingText;

    } else {

        button.disabled = false;

        if (button.dataset.originalText) {

            button.innerHTML = (
                button.dataset.originalText
            );

            delete button.dataset.originalText;
        }
    }
}


// ============================================================
// TOAST
// ============================================================

function showToast(
    message,
    type = "info"
) {

    let container = getElement(
        "toastContainer"
    );

    if (!container) {

        container = document.createElement(
            "div"
        );

        container.id = "toastContainer";

        container.className = "toast-container";

        document.body.appendChild(
            container
        );
    }

    const toast = document.createElement(
        "div"
    );

    toast.className = (
        `toast toast-${type}`
    );

    toast.textContent = message;

    container.appendChild(
        toast
    );

    requestAnimationFrame(
        () => {

            toast.classList.add(
                "visible"
            );
        }
    );

    setTimeout(
        () => {

            toast.classList.remove(
                "visible"
            );

            setTimeout(
                () => toast.remove(),
                250
            );
        },
        4500
    );
}


// ============================================================
// INICIALIZACIÓN
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    () => {

        loadCurrentAudit();

        /*
         * Los objetos File no pueden persistirse correctamente
         * en localStorage.
         *
         * Por seguridad, cuando recargamos el navegador
         * NO fingimos que los archivos siguen cargados.
         * Se conserva el memo, pero para una nueva extracción
         * hay que seleccionar los archivos nuevamente.
         */

        selectedFiles = [];

        state.files = [];

        state.sources = [];

        hydrateForm();

        bindGeneralFields();

        setupDropzone();

        renderFiles();

        renderExtraction();

        renderFindings();

        renderMemoPreview();

        updateFinalSummary();

        goToStep(
            state.currentStep || 1
        );

        // ----------------------------------------------
        // Botones encontrados por ID si existen
        // ----------------------------------------------

        const extractButton = (
            getElement("extractButton") ||
            getElement("analyzeButton")
        );

        if (extractButton) {

            extractButton.addEventListener(
                "click",
                extractInformation
            );
        }

        const exportButton = getElement(
            "exportExcelButton"
        );

        if (exportButton) {

            exportButton.addEventListener(
                "click",
                exportExcel
            );
        }

        const addFindingButton = getElement(
            "addFindingButton"
        );

        if (addFindingButton) {

            addFindingButton.addEventListener(
                "click",
                addBlankFinding
            );
        }

        const newAuditButton = getElement(
            "newAuditButton"
        );

        if (newAuditButton) {

            newAuditButton.addEventListener(
                "click",
                requestNewAudit
            );
        }
    }
);


// ============================================================
// FUNCIONES GLOBALES
//
// Necesarias porque algunos botones del HTML usan onclick.
// ============================================================

window.goToStep = goToStep;

window.nextStep = nextStep;

window.previousStep = previousStep;

window.extractInformation = extractInformation;

window.toggleExtractionIncluded = toggleExtractionIncluded;

window.updateExtractionText = updateExtractionText;

window.convertExtractionToFinding = convertExtractionToFinding;

window.addBlankFinding = addBlankFinding;

window.updateFinding = updateFinding;

window.deleteFinding = deleteFinding;

window.moveFinding = moveFinding;

window.removeFile = removeFile;

window.exportExcel = exportExcel;

window.improveField = improveField;

window.requestNewAudit = requestNewAudit;
