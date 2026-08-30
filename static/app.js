const state = {

    currentStep: 1,

    files: [],

    extracted: [],

    findings: [],

    sources: [],

    general: {
        title: "",
        area: "",
        process: "",
        period: "",
        auditor: "",
        objective: "",
        scope: ""
    }
};


/* =========================================================
   INICIO
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        loadState();

        bindGeneralFields();

        bindFiles();

        refreshGeneralFields();

        renderFiles();

        renderExtractions();

        renderFindings();

        goToStep(
            state.currentStep || 1
        );
    }
);


/* =========================================================
   DATOS GENERALES
========================================================= */

function bindGeneralFields() {

    const mapping = {
        auditTitle: "title",
        auditArea: "area",
        auditProcess: "process",
        auditPeriod: "period",
        auditAuditor: "auditor",
        auditObjective: "objective",
        auditScope: "scope"
    };


    Object.entries(
        mapping
    ).forEach(
        ([elementId, stateKey]) => {


            const element =
                document.getElementById(
                    elementId
                );


            if (!element) {
                return;
            }


            element.addEventListener(
                "input",
                () => {

                    state.general[
                        stateKey
                    ] = element.value;


                    saveState();
                }
            );


        }
    );
}


function refreshGeneralFields() {

    const mapping = {
        auditTitle: "title",
        auditArea: "area",
        auditProcess: "process",
        auditPeriod: "period",
        auditAuditor: "auditor",
        auditObjective: "objective",
        auditScope: "scope"
    };


    Object.entries(
        mapping
    ).forEach(
        ([elementId, stateKey]) => {


            const element =
                document.getElementById(
                    elementId
                );


            if (element) {

                element.value =
                    state.general[
                        stateKey
                    ]
                    || "";
            }


        }
    );
}


/* =========================================================
   MEJORAR TEXTO CON IA
========================================================= */

async function improveField(
    elementId,
    fieldType
) {

    const element =
        document.getElementById(
            elementId
        );


    if (!element) {
        return;
    }


    const originalText =
        element.value.trim();


    if (!originalText) {

        showToast(
            "Primero escribí una idea o texto para mejorar."
        );

        element.focus();

        return;
    }


    const buttons =
        document.querySelectorAll(
            ".ai-button"
        );


    buttons.forEach(
        button => {
            button.disabled = true;
        }
    );


    element.disabled = true;


    showToast(
        "Mejorando redacción con IA..."
    );


    try {

        const response =
            await fetch(
                "/improve-text",
                {

                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body: JSON.stringify({
                        text: originalText,
                        fieldType: fieldType
                    })

                }
            );


        const contentType =
            response.headers.get(
                "content-type"
            ) || "";


        if (
            !contentType.includes(
                "application/json"
            )
        ) {

            throw new Error(
                "El servidor devolvió una respuesta inválida."
            );
        }


        const data =
            await response.json();


        if (!response.ok) {

            throw new Error(
                data.error
                ||
                "No se pudo mejorar el texto."
            );
        }


        element.value =
            data.improved;


        element.dispatchEvent(
            new Event(
                "input",
                {
                    bubbles: true
                }
            )
        );


        showToast(
            "Texto mejorado. Podés editarlo libremente."
        );


    } catch (error) {

        console.error(
            error
        );


        showToast(
            error.message
            ||
            "No se pudo mejorar el texto."
        );


    } finally {

        element.disabled = false;


        buttons.forEach(
            button => {
                button.disabled = false;
            }
        );


        element.focus();
    }
}


/* =========================================================
   ARCHIVOS
========================================================= */

function bindFiles() {

    const fileInput =
        document.getElementById(
            "fileInput"
        );


    const dropZone =
        document.getElementById(
            "dropZone"
        );


    if (
        !fileInput
        ||
        !dropZone
    ) {

        return;
    }


    fileInput.addEventListener(
        "change",
        event => {

            addFiles(
                Array.from(
                    event.target.files
                )
            );


            fileInput.value = "";
        }
    );


    dropZone.addEventListener(
        "dragover",
        event => {

            event.preventDefault();

            dropZone.classList.add(
                "drag-over"
            );
        }
    );


    dropZone.addEventListener(
        "dragleave",
        () => {

            dropZone.classList.remove(
                "drag-over"
            );
        }
    );


    dropZone.addEventListener(
        "drop",
        event => {

            event.preventDefault();


            dropZone.classList.remove(
                "drag-over"
            );


            addFiles(
                Array.from(
                    event.dataTransfer.files
                )
            );
        }
    );
}


function addFiles(
    files
) {

    files.forEach(
        file => {


            const exists =
                state.files.some(
                    existing =>
                        existing.name === file.name
                        &&
                        existing.size === file.size
                );


            if (exists) {
                return;
            }


            state.files.push(
                file
            );


            const sourceAlreadyExists =
                state.sources.some(
                    source =>
                        source.name === file.name
                );


            if (!sourceAlreadyExists) {

                state.sources.push({

                    name:
                        file.name,

                    type:
                        getFileType(
                            file.name
                        ),

                    reference:
                        "",

                    description:
                        "Papel de trabajo cargado en Audit Memo Builder."

                });
            }


        }
    );


    renderFiles();

    saveState();
}


function removeFile(
    index
) {

    const file =
        state.files[
            index
        ];


    state.files.splice(
        index,
        1
    );


    if (file) {

        state.sources =
            state.sources.filter(
                source =>
                    source.name !== file.name
            );
    }


    renderFiles();

    saveState();
}


function renderFiles() {

    const container =
        document.getElementById(
            "fileList"
        );


    if (!container) {
        return;
    }


    container.innerHTML =
        "";


    state.files.forEach(
        (file, index) => {


            const chip =
                document.createElement(
                    "div"
                );


            chip.className =
                "file-chip";


            chip.innerHTML = `

                <span>
                    📄
                </span>

                <span>

                    ${escapeHtml(
                        file.name
                    )}

                    ·

                    ${formatBytes(
                        file.size
                    )}

                </span>

                <button
                    type="button"
                    onclick="removeFile(${index})"
                    title="Eliminar archivo"
                >
                    ×
                </button>
            `;


            container.appendChild(
                chip
            );


        }
    );
}


function getFileType(
    filename
) {

    const parts =
        filename.split(
            "."
        );


    if (
        parts.length < 2
    ) {

        return "Archivo";
    }


    const extension =
        parts.pop()
            .toLowerCase();


    const types = {
        xlsx: "Excel",
        csv: "CSV",
        docx: "Word",
        pdf: "PDF",
        txt: "Texto"
    };


    return types[
        extension
    ] || "Archivo";
}


function formatBytes(
    bytes
) {

    if (!bytes) {

        return "0 KB";
    }


    const mb =
        bytes
        /
        1024
        /
        1024;


    if (mb >= 1) {

        return `${mb.toFixed(1)} MB`;
    }


    return `${(
        bytes / 1024
    ).toFixed(0)} KB`;
}


/* =========================================================
   EXTRAER INFORMACIÓN
========================================================= */

async function extractInformation() {

    const freeTextElement =
        document.getElementById(
            "freeText"
        );


    const freeText =
        freeTextElement
        ?
        freeTextElement.value.trim()
        :
        "";


    if (
        state.files.length === 0
        &&
        !freeText
    ) {

        showToast(
            "Cargá al menos un papel de trabajo."
        );

        return;
    }


    const formData =
        new FormData();


    state.files.forEach(
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


    showToast(
        "Extrayendo información..."
    );


    try {

        const response =
            await fetch(
                "/extract",
                {
                    method: "POST",
                    body: formData
                }
            );


        const contentType =
            response.headers.get(
                "content-type"
            ) || "";


        if (
            !contentType.includes(
                "application/json"
            )
        ) {

            throw new Error(
                "El servidor devolvió una respuesta inválida. Revisá los logs de Render."
            );
        }


        const data =
            await response.json();


        if (!response.ok) {

            throw new Error(
                data.error
                ||
                "No se pudo extraer la información."
            );
        }


        state.extracted =
            Array.isArray(
                data.items
            )
            ?
            data.items.map(
                item => ({
                    ...item,
                    included:
                        Boolean(
                            item.included
                        ),
                    converted:
                        Boolean(
                            item.converted
                        )
                })
            )
            :
            [];


        saveState();


        renderExtractions();


        goToStep(
            2
        );


        const issueCount =
            (
                data.errors
                ?
                data.errors.length
                :
                0
            )
            +
            (
                data.warnings
                ?
                data.warnings.length
                :
                0
            );


        if (issueCount > 0) {

            showToast(
                `Extracción finalizada con ${issueCount} advertencia(s).`
            );

        } else {

            showToast(
                data.message
                ||
                "Información extraída."
            );
        }


    } catch (error) {

        console.error(
            error
        );


        showToast(
            error.message
            ||
            "Error al extraer información."
        );
    }
}


/* =========================================================
   INFORMACIÓN EXTRAÍDA
========================================================= */

function renderExtractions() {

    const list =
        document.getElementById(
            "extractionList"
        );


    const empty =
        document.getElementById(
            "extractionEmpty"
        );


    const count =
        document.getElementById(
            "extractedCount"
        );


    if (!list) {
        return;
    }


    list.innerHTML =
        "";


    if (count) {

        count.textContent =
            `${state.extracted.length} elemento(s)`;
    }


    if (
        state.extracted.length === 0
    ) {

        if (empty) {

            empty.style.display =
                "block";
        }


        return;
    }


    if (empty) {

        empty.style.display =
            "none";
    }


    state.extracted.forEach(
        item => {


            const card =
                document.createElement(
                    "div"
                );


            card.className =
                "extraction-card"
                +
                (
                    item.included
                    ?
                    " included"
                    :
                    ""
                );


            card.innerHTML = `

                <div class="extraction-check">

                    <input
                        type="checkbox"
                        ${item.included ? "checked" : ""}
                        onchange="
                            toggleExtraction(
                                '${item.id}',
                                this.checked
                            )
                        "
                    >

                </div>


                <div class="extraction-content">


                    <div class="extraction-top">


                        <span class="type-badge">

                            ${escapeHtml(
                                item.category
                                ||
                                "Contenido"
                            )}

                        </span>


                        <span class="origin">

                            ${escapeHtml(
                                item.filename
                                ||
                                ""
                            )}

                            ${
                                item.originName
                                ?
                                ` · ${escapeHtml(
                                    item.originName
                                )}`
                                :
                                ""
                            }

                        </span>


                    </div>


                    <textarea
                        class="extraction-text"
                        rows="3"
                        onchange="
                            updateExtractionText(
                                '${item.id}',
                                this.value
                            )
                        "
                    >${escapeHtml(
                        item.text
                        ||
                        ""
                    )}</textarea>


                    <div class="extraction-meta">


                        ${
                            item.originName
                            ?
                            `
                            <span class="meta-chip">
                                Solapa / origen:
                                ${escapeHtml(
                                    item.originName
                                )}
                            </span>
                            `
                            :
                            ""
                        }


                        ${
                            item.reference
                            ?
                            `
                            <span class="meta-chip">
                                ${escapeHtml(
                                    item.reference
                                )}
                            </span>
                            `
                            :
                            ""
                        }


                        ${
                            item.keyword
                            ?
                            `
                            <span class="meta-chip">

                                Detectado por:
                                ${escapeHtml(
                                    item.keyword
                                )}

                            </span>
                            `
                            :
                            ""
                        }


                    </div>


                </div>
            `;


            list.appendChild(
                card
            );


        }
    );
}


function toggleExtraction(
    id,
    included
) {

    const item =
        state.extracted.find(
            element =>
                element.id === id
        );


    if (!item) {
        return;
    }


    item.included =
        included;


    renderExtractions();

    saveState();
}


function updateExtractionText(
    id,
    text
) {

    const item =
        state.extracted.find(
            element =>
                element.id === id
        );


    if (!item) {
        return;
    }


    item.text =
        text;


    saveState();
}


/* =========================================================
   CONVERTIR A HALLAZGOS
========================================================= */

function convertSelectedToFindings() {

    const selected =
        state.extracted.filter(
            item =>
                item.included
                &&
                !item.converted
        );


    selected.forEach(
        item => {


            state.findings.push({

                id:
                    generateId(),

                title:
                    createSuggestedTitle(
                        item
                    ),

                situation:
                    item.text
                    ||
                    "",

                risk:
                    "",

                proposal:
                    "",

                responsibleArea:
                    "",

                actionOwner:
                    "",

                severity:
                    "",

                status:
                    "Pendiente",

                targetDate:
                    "",

                quantitativeBasis:
                    "",

                sourceFile:
                    item.filename
                    ||
                    "",

                sourceLocation:
                    item.originName
                    ||
                    "",

                evidence:
                    item.reference
                    ||
                    "",

                ticket:
                    "",

                followUp:
                    "",

                extractionId:
                    item.id
            });


            item.converted =
                true;


        }
    );


    saveState();

    renderFindings();

    goToStep(
        3
    );


    if (
        selected.length === 0
        &&
        state.findings.length === 0
    ) {

        showToast(
            "Seleccioná información o agregá un hallazgo manual."
        );
    }
}


function createSuggestedTitle(
    item
) {

    let text =
        String(
            item.text
            ||
            ""
        )
        .replace(
            /\|/g,
            " "
        )
        .replace(
            /\s+/g,
            " "
        )
        .trim();


    if (!text) {

        return (
            item.category
            ||
            "Hallazgo"
        );
    }


    if (
        text.length > 90
    ) {

        text =
            text.substring(
                0,
                87
            )
            +
            "...";
    }


    return text;
}


/* =========================================================
   HALLAZGOS
========================================================= */

function addFinding() {

    state.findings.push({

        id:
            generateId(),

        title:
            "",

        situation:
            "",

        risk:
            "",

        proposal:
            "",

        responsibleArea:
            "",

        actionOwner:
            "",

        severity:
            "",

        status:
            "Pendiente",

        targetDate:
            "",

        quantitativeBasis:
            "",

        sourceFile:
            "",

        sourceLocation:
            "",

        evidence:
            "",

        ticket:
            "",

        followUp:
            "",

        extractionId:
            ""

    });


    renderFindings();

    saveState();
}


function deleteFinding(
    id
) {

    const finding =
        state.findings.find(
            item =>
                item.id === id
        );


    if (
        finding
        &&
        finding.extractionId
    ) {

        const extraction =
            state.extracted.find(
                item =>
                    item.id
                    ===
                    finding.extractionId
            );


        if (extraction) {

            extraction.converted =
                false;
        }
    }


    state.findings =
        state.findings.filter(
            finding =>
                finding.id !== id
        );


    renderFindings();

    saveState();
}


function moveFinding(
    id,
    direction
) {

    const index =
        state.findings.findIndex(
            finding =>
                finding.id === id
        );


    if (
        index === -1
    ) {

        return;
    }


    const newIndex =
        index + direction;


    if (
        newIndex < 0
        ||
        newIndex >= state.findings.length
    ) {

        return;
    }


    const [finding] =
        state.findings.splice(
            index,
            1
        );


    state.findings.splice(
        newIndex,
        0,
        finding
    );


    renderFindings();

    saveState();
}


function updateFinding(
    id,
    field,
    value
) {

    const finding =
        state.findings.find(
            item =>
                item.id === id
        );


    if (!finding) {
        return;
    }


    finding[field] =
        value;


    saveState();
}


function setSeverity(
    id,
    severity
) {

    updateFinding(
        id,
        "severity",
        severity
    );


    renderFindings();
}


function renderFindings() {

    const list =
        document.getElementById(
            "findingsList"
        );


    const empty =
        document.getElementById(
            "findingsEmpty"
        );


    if (!list) {
        return;
    }


    list.innerHTML =
        "";


    if (
        state.findings.length === 0
    ) {

        if (empty) {

            empty.style.display =
                "block";
        }


        return;
    }


    if (empty) {

        empty.style.display =
            "none";
    }


    state.findings.forEach(
        (finding, index) => {


            const card =
                document.createElement(
                    "article"
                );


            card.className =
                "finding-card";


            card.innerHTML = `

                <div class="finding-header">


                    <div>


                        <div class="finding-number">

                            HALLAZGO
                            ${String(
                                index + 1
                            ).padStart(
                                2,
                                "0"
                            )}

                        </div>


                        <div class="finding-origin">

                            ${
                                finding.sourceFile
                                ?
                                `
                                Origen:
                                ${escapeHtml(
                                    finding.sourceFile
                                )}
                                `
                                :
                                "Hallazgo manual"
                            }

                            ${
                                finding.sourceLocation
                                ?
                                `
                                · Solapa:
                                ${escapeHtml(
                                    finding.sourceLocation
                                )}
                                `
                                :
                                ""
                            }

                        </div>


                    </div>


                    <div class="finding-actions">


                        <button
                            class="finding-delete"
                            type="button"
                            onclick="
                                moveFinding(
                                    '${finding.id}',
                                    -1
                                )
                            "
                            title="Mover arriba"
                            ${index === 0 ? "disabled" : ""}
                        >
                            ↑
                        </button>


                        <button
                            class="finding-delete"
                            type="button"
                            onclick="
                                moveFinding(
                                    '${finding.id}',
                                    1
                                )
                            "
                            title="Mover abajo"
                            ${
                                index ===
                                state.findings.length - 1
                                ?
                                "disabled"
                                :
                                ""
                            }
                        >
                            ↓
                        </button>


                        <button
                            class="finding-delete"
                            type="button"
                            onclick="
                                deleteFinding(
                                    '${finding.id}'
                                )
                            "
                            title="Eliminar"
                        >
                            ×
                        </button>


                    </div>


                </div>


                <div class="finding-body">


                    <div class="field">


                        <label>
                            Título del hallazgo
                        </label>


                        <input
                            value="${escapeHtml(
                                finding.title
                            )}"
                            oninput="
                                updateFinding(
                                    '${finding.id}',
                                    'title',
                                    this.value
                                )
                            "
                        >


                    </div>


                    <div class="field">


                        <label>
                            Situación observada
                        </label>


                        <textarea
                            rows="5"
                            oninput="
                                updateFinding(
                                    '${finding.id}',
                                    'situation',
                                    this.value
                                )
                            "
                        >${escapeHtml(
                            finding.situation
                        )}</textarea>


                    </div>


                    <div class="field">


                        <label>
                            Riesgo
                        </label>


                        <textarea
                            rows="3"
                            oninput="
                                updateFinding(
                                    '${finding.id}',
                                    'risk',
                                    this.value
                                )
                            "
                        >${escapeHtml(
                            finding.risk
                        )}</textarea>


                    </div>


                    <div class="field">


                        <label>
                            Propuesta de mejora
                        </label>


                        <textarea
                            rows="3"
                            oninput="
                                updateFinding(
                                    '${finding.id}',
                                    'proposal',
                                    this.value
                                )
                            "
                        >${escapeHtml(
                            finding.proposal
                        )}</textarea>


                    </div>


                    <div class="finding-row">


                        <div class="field">

                            <label>
                                Área responsable
                            </label>

                            <input
                                value="${escapeHtml(
                                    finding.responsibleArea
                                )}"
                                oninput="
                                    updateFinding(
                                        '${finding.id}',
                                        'responsibleArea',
                                        this.value
                                    )
                                "
                            >

                        </div>


                        <div class="field">

                            <label>
                                Estado
                            </label>

                            <select
                                onchange="
                                    updateFinding(
                                        '${finding.id}',
                                        'status',
                                        this.value
                                    )
                                "
                            >

                                ${renderStatusOptions(
                                    finding.status
                                )}

                            </select>

                        </div>


                    </div>


                    <div class="field">


                        <label>
                            Criticidad
                        </label>


                        <div class="severity-selector">


                            ${severityButton(
                                finding,
                                "Alta",
                                "high"
                            )}


                            ${severityButton(
                                finding,
                                "Media",
                                "medium"
                            )}


                            ${severityButton(
                                finding,
                                "Baja",
                                "low"
                            )}


                        </div>


                    </div>


                    <details class="details-box">


                        <summary>
                            Más detalles
                        </summary>


                        <div class="details-grid">


                            <div class="field">

                                <label>
                                    Responsable del plan
                                </label>

                                <input
                                    value="${escapeHtml(
                                        finding.actionOwner
                                    )}"
                                    oninput="
                                        updateFinding(
                                            '${finding.id}',
                                            'actionOwner',
                                            this.value
                                        )
                                    "
                                >

                            </div>


                            <div class="field">

                                <label>
                                    Fecha compromiso
                                </label>

                                <input
                                    type="date"
                                    value="${escapeHtml(
                                        finding.targetDate
                                    )}"
                                    onchange="
                                        updateFinding(
                                            '${finding.id}',
                                            'targetDate',
                                            this.value
                                        )
                                    "
                                >

                            </div>


                            <div class="field">

                                <label>
                                    Base cuantitativa
                                </label>

                                <input
                                    value="${escapeHtml(
                                        finding.quantitativeBasis
                                    )}"
                                    oninput="
                                        updateFinding(
                                            '${finding.id}',
                                            'quantitativeBasis',
                                            this.value
                                        )
                                    "
                                >

                            </div>


                            <div class="field">

                                <label>
                                    Ticket / referencia
                                </label>

                                <input
                                    value="${escapeHtml(
                                        finding.ticket
                                    )}"
                                    oninput="
                                        updateFinding(
                                            '${finding.id}',
                                            'ticket',
                                            this.value
                                        )
                                    "
                                >

                            </div>


                            <div class="field">

                                <label>
                                    Archivo de origen
                                </label>

                                <input
                                    value="${escapeHtml(
                                        finding.sourceFile
                                    )}"
                                    oninput="
                                        updateFinding(
                                            '${finding.id}',
                                            'sourceFile',
                                            this.value
                                        )
                                    "
                                >

                            </div>


                            <div class="field">

                                <label>
                                    Solapa / origen
                                </label>

                                <input
                                    value="${escapeHtml(
                                        finding.sourceLocation
                                    )}"
                                    oninput="
                                        updateFinding(
                                            '${finding.id}',
                                            'sourceLocation',
                                            this.value
                                        )
                                    "
                                >

                            </div>


                            <div class="field">

                                <label>
                                    Evidencia / referencia
                                </label>

                                <input
                                    value="${escapeHtml(
                                        finding.evidence
                                    )}"
                                    oninput="
                                        updateFinding(
                                            '${finding.id}',
                                            'evidence',
                                            this.value
                                        )
                                    "
                                >

                            </div>


                            <div class="field">

                                <label>
                                    Seguimiento
                                </label>

                                <textarea
                                    rows="2"
                                    oninput="
                                        updateFinding(
                                            '${finding.id}',
                                            'followUp',
                                            this.value
                                        )
                                    "
                                >${escapeHtml(
                                    finding.followUp
                                )}</textarea>

                            </div>


                        </div>


                    </details>


                </div>
            `;


            list.appendChild(
                card
            );


        }
    );
}


function severityButton(
    finding,
    label,
    cssClass
) {

    const selected =
        finding.severity === label
        ?
        "selected"
        :
        "";


    let icon =
        "🟢";


    if (
        label === "Alta"
    ) {

        icon =
            "🔴";
    }


    if (
        label === "Media"
    ) {

        icon =
            "🟡";
    }


    return `

        <button
            type="button"
            class="
                severity-option
                ${cssClass}
                ${selected}
            "
            onclick="
                setSeverity(
                    '${finding.id}',
                    '${label}'
                )
            "
        >

            ${icon}
            ${label}

        </button>
    `;
}


function renderStatusOptions(
    selected
) {

    const options = [
        "Pendiente",
        "En curso",
        "Implementado",
        "Cerrado"
    ];


    return options.map(
        option => `

            <option
                value="${option}"
                ${
                    option === selected
                    ?
                    "selected"
                    :
                    ""
                }
            >
                ${option}
            </option>

        `
    ).join(
        ""
    );
}


/* =========================================================
   NAVEGACIÓN
========================================================= */

function goToStep(
    step
) {

    const target =
        document.getElementById(
            `step${step}`
        );


    if (!target) {
        return;
    }


    state.currentStep =
        step;


    document
        .querySelectorAll(
            ".page-step"
        )
        .forEach(
            section => {

                section.classList.remove(
                    "active"
                );
            }
        );


    target.classList.add(
        "active"
    );


    document
        .querySelectorAll(
            ".step-item"
        )
        .forEach(
            item => {

                item.classList.toggle(
                    "active",
                    Number(
                        item.dataset.step
                    ) === step
                );
            }
        );


    if (
        step === 2
    ) {

        renderExtractions();
    }


    if (
        step === 3
    ) {

        renderFindings();
    }


    if (
        step === 4
    ) {

        renderMemo();
    }


    if (
        step === 5
    ) {

        renderValidation();
    }


    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });


    saveState();
}


/* =========================================================
   MEMO
========================================================= */

function renderMemo() {

    const container =
        document.getElementById(
            "memoPreview"
        );


    if (!container) {
        return;
    }


    const g =
        state.general;


    const findingsHtml =
        state.findings.map(
            (finding, index) => {


                let severityClass =
                    "";


                if (
                    finding.severity === "Alta"
                ) {

                    severityClass =
                        "high";
                }


                if (
                    finding.severity === "Media"
                ) {

                    severityClass =
                        "medium";
                }


                if (
                    finding.severity === "Baja"
                ) {

                    severityClass =
                        "low";
                }


                return `

                    <section class="memo-finding">


                        <div class="finding-number">

                            HALLAZGO
                            ${String(
                                index + 1
                            ).padStart(
                                2,
                                "0"
                            )}

                        </div>


                        <h3 class="memo-finding-title">

                            ${escapeHtml(
                                finding.title
                                ||
                                "Sin título"
                            )}

                        </h3>


                        <div class="memo-tags">


                            ${
                                finding.severity
                                ?
                                `
                                <span
                                    class="
                                        memo-tag
                                        ${severityClass}
                                    "
                                >

                                    Criticidad:
                                    ${escapeHtml(
                                        finding.severity
                                    )}

                                </span>
                                `
                                :
                                ""
                            }


                            ${
                                finding.responsibleArea
                                ?
                                `
                                <span class="memo-tag">

                                    Área responsable:
                                    ${escapeHtml(
                                        finding.responsibleArea
                                    )}

                                </span>
                                `
                                :
                                ""
                            }


                            ${
                                finding.status
                                ?
                                `
                                <span class="memo-tag">

                                    Estado:
                                    ${escapeHtml(
                                        finding.status
                                    )}

                                </span>
                                `
                                :
                                ""
                            }


                        </div>


                        <div class="memo-section">

                            <h3>
                                Situación observada
                            </h3>

                            <p>
                                ${escapeHtml(
                                    finding.situation
                                    ||
                                    "Pendiente de completar."
                                )}
                            </p>

                        </div>


                        <div class="memo-section">

                            <h3>
                                Riesgo
                            </h3>

                            <p>
                                ${escapeHtml(
                                    finding.risk
                                    ||
                                    "Pendiente de completar."
                                )}
                            </p>

                        </div>


                        <div class="memo-section">

                            <h3>
                                Propuesta de mejora
                            </h3>

                            <p>
                                ${escapeHtml(
                                    finding.proposal
                                    ||
                                    "Pendiente de completar."
                                )}
                            </p>

                        </div>


                        ${
                            finding.sourceFile
                            ||
                            finding.sourceLocation
                            ||
                            finding.evidence
                            ?
                            `
                            <div class="memo-section">

                                <h3>
                                    Trazabilidad
                                </h3>

                                <p>

                                    ${
                                        finding.sourceFile
                                        ?
                                        `
                                        <strong>
                                            Archivo:
                                        </strong>

                                        ${escapeHtml(
                                            finding.sourceFile
                                        )}

                                        <br>
                                        `
                                        :
                                        ""
                                    }

                                    ${
                                        finding.sourceLocation
                                        ?
                                        `
                                        <strong>
                                            Solapa / origen:
                                        </strong>

                                        ${escapeHtml(
                                            finding.sourceLocation
                                        )}

                                        <br>
                                        `
                                        :
                                        ""
                                    }

                                    ${
                                        finding.evidence
                                        ?
                                        `
                                        <strong>
                                            Referencia:
                                        </strong>

                                        ${escapeHtml(
                                            finding.evidence
                                        )}
                                        `
                                        :
                                        ""
                                    }

                                </p>

                            </div>
                            `
                            :
                            ""
                        }


                    </section>
                `;


            }
        ).join(
            ""
        );


    container.innerHTML = `


        <div class="memo-company">
            AUDITORÍA INTERNA
        </div>


        <h1 class="memo-title">

            ${escapeHtml(
                g.title
                ||
                "Memo de Auditoría"
            )}

        </h1>


        <div class="memo-meta">


            <div>

                <strong>
                    Área:
                </strong>

                ${escapeHtml(
                    g.area
                    ||
                    "-"
                )}

            </div>


            <div>

                <strong>
                    Proceso:
                </strong>

                ${escapeHtml(
                    g.process
                    ||
                    "-"
                )}

            </div>


            <div>

                <strong>
                    Período:
                </strong>

                ${escapeHtml(
                    g.period
                    ||
                    "-"
                )}

            </div>


            <div>

                <strong>
                    Auditor:
                </strong>

                ${escapeHtml(
                    g.auditor
                    ||
                    "-"
                )}

            </div>


        </div>


        <section class="memo-section">

            <h3>
                Objetivo
            </h3>

            <p>
                ${escapeHtml(
                    g.objective
                    ||
                    "Pendiente de completar."
                )}
            </p>

        </section>


        <section class="memo-section">

            <h3>
                Alcance
            </h3>

            <p>
                ${escapeHtml(
                    g.scope
                    ||
                    "Pendiente de completar."
                )}
            </p>

        </section>


        ${
            findingsHtml
            ||
            `
            <section class="memo-section">

                <h3>
                    Hallazgos
                </h3>

                <p>
                    No se incorporaron hallazgos.
                </p>

            </section>
            `
        }
    `;
}


/* =========================================================
   VALIDACIÓN FINAL
========================================================= */

function renderValidation() {

    const container =
        document.getElementById(
            "validationList"
        );


    if (!container) {
        return;
    }


    const checks =
        [];


    checks.push({

        ok:
            Boolean(
                state.general.title
            ),

        text:
            state.general.title
            ?
            "Nombre de auditoría informado."
            :
            "Falta informar el nombre de la auditoría."

    });


    checks.push({

        ok:
            Boolean(
                state.general.objective
            ),

        text:
            state.general.objective
            ?
            "Objetivo informado."
            :
            "Falta completar el objetivo."

    });


    checks.push({

        ok:
            Boolean(
                state.general.scope
            ),

        text:
            state.general.scope
            ?
            "Alcance informado."
            :
            "Falta completar el alcance."

    });


    checks.push({

        ok:
            state.sources.length > 0,

        text:
            state.sources.length > 0
            ?
            `${state.sources.length} fuente(s) registrada(s).`
            :
            "No hay fuentes registradas."

    });


    checks.push({

        ok:
            state.findings.length > 0,

        text:
            state.findings.length > 0
            ?
            `${state.findings.length} hallazgo(s) incluido(s).`
            :
            "No hay hallazgos incluidos."

    });


    const incompleteFindings =
        state.findings.filter(
            finding =>
                !finding.title
                ||
                !finding.situation
                ||
                !finding.risk
                ||
                !finding.proposal
                ||
                !finding.responsibleArea
                ||
                !finding.severity
        );


    checks.push({

        ok:
            incompleteFindings.length === 0,

        text:
            incompleteFindings.length === 0
            ?
            "Todos los hallazgos contienen los campos principales."
            :
            `${incompleteFindings.length} hallazgo(s) tienen campos pendientes.`

    });


    const withoutDate =
        state.findings.filter(
            finding =>
                !finding.targetDate
        );


    checks.push({

        ok:
            withoutDate.length === 0,

        text:
            withoutDate.length === 0
            ?
            "Todos los hallazgos tienen fecha compromiso."
            :
            `${withoutDate.length} hallazgo(s) sin fecha compromiso.`

    });


    container.innerHTML =
        checks.map(
            check => `

                <div
                    class="
                        validation-item
                        ${
                            check.ok
                            ?
                            "ok"
                            :
                            "warning"
                        }
                    "
                >

                    <strong>
                        ${
                            check.ok
                            ?
                            "✓"
                            :
                            "!"
                        }
                    </strong>

                    <span>
                        ${escapeHtml(
                            check.text
                        )}
                    </span>

                </div>

            `
        ).join(
            ""
        );
}


/* =========================================================
   EXPORTAR
========================================================= */

async function exportExcel() {

    const memo = {

        general:
            state.general,

        findings:
            state.findings,

        sources:
            state.sources,

        extracted:
            state.extracted
    };


    try {

        showToast(
            "Generando memo..."
        );


        const response =
            await fetch(
                "/export-excel",
                {

                    method:
                        "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({
                            memo
                        })

                }
            );


        if (!response.ok) {

            throw new Error(
                "No se pudo generar el Excel."
            );
        }


        const blob =
            await response.blob();


        const url =
            window.URL.createObjectURL(
                blob
            );


        const link =
            document.createElement(
                "a"
            );


        link.href =
            url;


        link.download =
            createExportFilename();


        document.body.appendChild(
            link
        );


        link.click();


        link.remove();


        window.URL.revokeObjectURL(
            url
        );


        showToast(
            "Memo exportado correctamente."
        );


    } catch (error) {

        console.error(
            error
        );


        showToast(
            error.message
        );
    }
}


function createExportFilename() {

    const title =
        state.general.title
        ||
        "Audit_Memo";


    const safeTitle =
        title
        .replace(
            /[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ_-]/g,
            "_"
        )
        .replace(
            /_+/g,
            "_"
        );


    return `${safeTitle}.xlsx`;
}


/* =========================================================
   LOCAL STORAGE
========================================================= */

function saveState() {

    const persistable = {

        currentStep:
            state.currentStep,

        general:
            state.general,

        extracted:
            state.extracted,

        findings:
            state.findings,

        sources:
            state.sources
    };


    localStorage.setItem(
        "auditMemoBuilderState",
        JSON.stringify(
            persistable
        )
    );
}


function loadState() {

    try {

        const saved =
            localStorage.getItem(
                "auditMemoBuilderState"
            );


        if (!saved) {
            return;
        }


        const parsed =
            JSON.parse(
                saved
            );


        state.currentStep =
            parsed.currentStep
            ||
            1;


        state.general =
            parsed.general
            ||
            state.general;


        state.extracted =
            parsed.extracted
            ||
            [];


        state.findings =
            parsed.findings
            ||
            [];


        state.sources =
            parsed.sources
            ||
            [];


        state.files =
            [];


    } catch (error) {

        console.error(
            "No se pudo recuperar el trabajo.",
            error
        );
    }
}


/* =========================================================
   UTILIDADES
========================================================= */

function generateId() {

    if (
        window.crypto
        &&
        crypto.randomUUID
    ) {

        return crypto.randomUUID();
    }


    return (
        Date.now().toString(
            36
        )
        +
        Math.random()
            .toString(
                36
            )
            .substring(
                2
            )
    );
}


function escapeHtml(
    value
) {

    if (
        value === null
        ||
        value === undefined
    ) {

        return "";
    }


    return String(
        value
    )
    .replace(
        /&/g,
        "&amp;"
    )
    .replace(
        /</g,
        "&lt;"
    )
    .replace(
        />/g,
        "&gt;"
    )
    .replace(
        /"/g,
        "&quot;"
    )
    .replace(
        /'/g,
        "&#039;"
    );
}


function showToast(
    message
) {

    const toast =
        document.getElementById(
            "toast"
        );


    if (!toast) {
        return;
    }


    toast.textContent =
        message;


    toast.classList.add(
        "show"
    );


    clearTimeout(
        window.toastTimeout
    );


    window.toastTimeout =
        setTimeout(
            () => {

                toast.classList.remove(
                    "show"
                );

            },
            3500
        );
}
