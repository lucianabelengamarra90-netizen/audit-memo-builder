const state = {
    files: [],
    facts: [],
    tasks: [],
    results: [],
    findings: [],
    sources: [],
    memo: null
};

const $ = (id) => document.getElementById(id);


// =========================================================
// AUXILIARES
// =========================================================

function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        "'": "&#39;",
        '"': "&quot;"
    }[char]));
}

function escapeAttribute(value) {
    return escapeHtml(value);
}

function nl2br(value) {
    return escapeHtml(value).replace(/\n/g, "<br>");
}

function notify(message, type = "success") {
    const element = document.createElement("div");

    element.className = `toast ${type}`;
    element.textContent = message;

    document.body.appendChild(element);

    setTimeout(() => element.remove(), 4200);
}

function fileSize(size) {
    if (size < 1024) {
        return `${size} B`;
    }

    if (size < 1024 * 1024) {
        return `${(size / 1024).toFixed(1)} KB`;
    }

    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function generateId() {
    if (
        window.crypto &&
        typeof window.crypto.randomUUID === "function"
    ) {
        return window.crypto.randomUUID();
    }

    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}


// =========================================================
// DASHBOARD
// =========================================================

function updateDashboard() {
    const accepted = state.facts.filter(
        (fact) => fact.status === "accepted"
    ).length;

    const criticality = {
        high: 0,
        medium: 0,
        low: 0
    };

    state.findings.forEach((finding) => {
        if (finding.criticidad) {
            criticality[finding.criticidad] =
                (criticality[finding.criticidad] || 0) + 1;
        }
    });

    const values = {
        metricFiles: state.files.length,
        metricFacts: state.facts.length,
        metricAccepted: accepted,
        metricTasks: state.tasks.length,
        metricFindings: state.findings.length,
        metricHigh: criticality.high,
        metricMedium: criticality.medium,
        metricLow: criticality.low
    };

    Object.entries(values).forEach(([id, value]) => {
        if ($(id)) {
            $(id).textContent = value;
        }
    });
}


// =========================================================
// NAVEGACIÓN
// =========================================================

function goToStep(step) {
    if (step === 6 && !validateBeforeMemo()) {
        return;
    }

    document
        .querySelectorAll(".step-section")
        .forEach((section) => {
            section.classList.remove("active");
        });

    document
        .querySelectorAll(".step-nav-item")
        .forEach((item) => {
            item.classList.toggle(
                "active",
                Number(item.dataset.step) === step
            );
        });

    const target = $(`step-${step}`);

    if (target) {
        target.classList.add("active");
    }

    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
}


// =========================================================
// ARCHIVOS
// =========================================================

function initializeUpload() {
    const dropZone = $("dropZone");
    const fileInput = $("fileInput");

    if (!dropZone || !fileInput) {
        console.error("No se encontró dropZone o fileInput.");
        return;
    }

    fileInput.addEventListener("change", function () {
        const selectedFiles = Array.from(fileInput.files || []);

        if (!selectedFiles.length) {
            return;
        }

        addFiles(selectedFiles);

        fileInput.value = "";
    });

    dropZone.addEventListener("dragover", function (event) {
        event.preventDefault();

        dropZone.classList.add("dragging");
    });

    dropZone.addEventListener("dragleave", function () {
        dropZone.classList.remove("dragging");
    });

    dropZone.addEventListener("drop", function (event) {
        event.preventDefault();

        dropZone.classList.remove("dragging");

        const droppedFiles = Array.from(
            event.dataTransfer?.files || []
        );

        if (droppedFiles.length) {
            addFiles(droppedFiles);
        }
    });
}

function addFiles(files) {
    const validExtensions = [
        "xlsx",
        "xls",
        "docx",
        "pdf",
        "csv",
        "txt"
    ];

    let rejected = 0;
    let added = 0;
    let duplicated = 0;

    files.forEach(function (file) {
        const fileName = file.name || "";

        const extension = fileName
            .split(".")
            .pop()
            .toLowerCase();

        if (!validExtensions.includes(extension)) {
            rejected++;
            return;
        }

        const alreadyExists = state.files.some(function (item) {
            return (
                item.name === file.name &&
                item.size === file.size
            );
        });

        if (alreadyExists) {
            duplicated++;
            return;
        }

        state.files.push({
            id: generateId(),
            file: file,
            name: file.name,
            type: extension,
            size: file.size
        });

        added++;

        const sourceExists = state.sources.some(function (source) {
            return source.name === file.name;
        });

        if (!sourceExists) {
            state.sources.push({
                id: generateId(),
                name: file.name,
                type: sourceTypeFromExtension(extension),
                reference: file.name,
                description: ""
            });
        }
    });

    renderFiles();
    renderSources();
    updateDashboard();

    if (added > 0) {
        notify(
            `${added} archivo(s) cargado(s) correctamente.`
        );
    }

    if (duplicated > 0) {
        notify(
            `${duplicated} archivo(s) ya estaban cargados.`,
            "warning"
        );
    }

    if (rejected > 0) {
        notify(
            `${rejected} archivo(s) tienen un formato no admitido.`,
            "warning"
        );
    }
}

function sourceTypeFromExtension(extension) {
    if (
        ["xlsx", "xls", "csv"].includes(extension)
    ) {
        return "Excel / Datos";
    }

    if (extension === "pdf") {
        return "PDF";
    }

    if (extension === "docx") {
        return "Word";
    }

    if (extension === "txt") {
        return "Texto";
    }

    return "Otro";
}

function renderFiles() {
    const container = $("filesContainer");

    if (!container) {
        return;
    }

    if (!state.files.length) {
        container.innerHTML =
            '<p class="section-description">Sin archivos cargados.</p>';

        return;
    }

    container.innerHTML = state.files
        .map((item) => `
            <div class="file-item">
                <div class="file-meta">
                    <span class="file-type">
                        ${escapeHtml(item.type)}
                    </span>

                    <strong>
                        ${escapeHtml(item.name)}
                    </strong>

                    <span class="file-size">
                        ${fileSize(item.size)}
                    </span>
                </div>

                <button
                    class="btn btn-danger btn-sm"
                    type="button"
                    onclick="removeFile('${item.id}')"
                >
                    Eliminar
                </button>
            </div>
        `)
        .join("");
}

function removeFile(id) {
    const selected = state.files.find(
        (file) => file.id === id
    );

    state.files = state.files.filter(
        (file) => file.id !== id
    );

    if (selected) {
        state.sources = state.sources.filter(
            (source) =>
                !(
                    source.name === selected.name &&
                    source.reference === selected.name
                )
        );
    }

    renderFiles();
    renderSources();
    updateDashboard();
}


// =========================================================
// OBJETIVOS
// =========================================================

function addObjective(text = "") {
    const id = generateId();

    $("objectivesContainer").insertAdjacentHTML(
        "beforeend",
        `
        <div
            class="objective-row"
            data-id="${id}"
        >
            <input
                type="text"
                value="${escapeAttribute(text)}"
                placeholder="Describa el objetivo de auditoría"
            >

            <button
                class="btn btn-secondary reorder-btn"
                type="button"
                onclick="moveObjective('${id}', -1)"
                title="Subir"
            >
                ↑
            </button>

            <button
                class="btn btn-secondary reorder-btn"
                type="button"
                onclick="moveObjective('${id}', 1)"
                title="Bajar"
            >
                ↓
            </button>

            <button
                class="btn btn-danger btn-sm"
                type="button"
                onclick="removeObjective('${id}')"
            >
                Eliminar
            </button>
        </div>
        `
    );
}

function removeObjective(id) {
    document
        .querySelector(
            `.objective-row[data-id="${id}"]`
        )
        ?.remove();
}

function moveObjective(id, direction) {
    const row = document.querySelector(
        `.objective-row[data-id="${id}"]`
    );

    const container = $("objectivesContainer");

    if (!row || !container) {
        return;
    }

    if (
        direction < 0 &&
        row.previousElementSibling
    ) {
        container.insertBefore(
            row,
            row.previousElementSibling
        );
    }

    if (
        direction > 0 &&
        row.nextElementSibling
    ) {
        container.insertBefore(
            row.nextElementSibling,
            row
        );
    }
}

function getObjectives() {
    return [
        ...document.querySelectorAll(
            ".objective-row input"
        )
    ]
        .map((input) => input.value.trim())
        .filter(Boolean);
}


// =========================================================
// FUENTES DINÁMICAS
// =========================================================

function addSource(data = {}) {
    state.sources.push({
        id: data.id || generateId(),
        name: data.name || "",
        type: data.type || "",
        reference: data.reference || "",
        description: data.description || ""
    });

    renderSources();
}

function renderSources() {
    const container = $("sourcesContainer");

    if (!container) {
        return;
    }

    if (!state.sources.length) {
        container.innerHTML =
            '<p class="section-description">Todavía no agregaste fuentes de información.</p>';

        return;
    }

    const types = [
        "",
        "SAP",
        "Excel / Datos",
        "PDF",
        "Word",
        "Mail",
        "Entrevista / Reunión",
        "Power BI",
        "Jira / Ticket",
        "Procedimiento",
        "Contrato",
        "Confirmación externa",
        "Texto",
        "Otro"
    ];

    container.innerHTML = state.sources
        .map(
            (
                source,
                index
            ) => `
            <div
                class="source-row"
                data-id="${source.id}"
            >
                <div class="source-row-header">
                    <strong>
                        Fuente ${String(index + 1).padStart(2, "0")}
                    </strong>

                    <button
                        class="btn btn-danger btn-sm"
                        type="button"
                        onclick="deleteSource('${source.id}')"
                    >
                        Eliminar
                    </button>
                </div>

                <div class="source-grid">
                    <div class="form-group">
                        <label>
                            Nombre de la fuente *
                        </label>

                        <input
                            data-field="name"
                            value="${escapeAttribute(source.name)}"
                            placeholder="Ej. FBL1N - Cuenta corriente de proveedores"
                            oninput="syncSource('${source.id}', this)"
                        >
                    </div>

                    <div class="form-group">
                        <label>
                            Tipo
                        </label>

                        <select
                            data-field="type"
                            onchange="syncSource('${source.id}', this)"
                        >
                            ${types
                                .map(
                                    (type) => `
                                    <option
                                        value="${escapeAttribute(type)}"
                                        ${source.type === type ? "selected" : ""}
                                    >
                                        ${escapeHtml(type || "Seleccione")}
                                    </option>
                                    `
                                )
                                .join("")
                            }
                        </select>
                    </div>

                    <div class="form-group">
                        <label>
                            Referencia
                        </label>

                        <input
                            data-field="reference"
                            value="${escapeAttribute(source.reference)}"
                            placeholder="Hoja, transacción, ticket, reporte..."
                            oninput="syncSource('${source.id}', this)"
                        >
                    </div>

                    <div class="form-group">
                        <label>
                            Descripción / alcance
                        </label>

                        <input
                            data-field="description"
                            value="${escapeAttribute(source.description)}"
                            placeholder="Qué información aporta esta fuente"
                            oninput="syncSource('${source.id}', this)"
                        >
                    </div>
                </div>
            </div>
            `
        )
        .join("");
}

function syncSource(id, element) {
    const source = state.sources.find(
        (item) => item.id === id
    );

    if (source) {
        source[element.dataset.field] =
            element.value;
    }
}

function deleteSource(id) {
    state.sources = state.sources.filter(
        (source) => source.id !== id
    );

    renderSources();
}

function getSources() {
    return state.sources
        .filter(
            (source) =>
                (source.name || "").trim()
        )
        .map(
            (source) => ({
                name:
                    (source.name || "").trim(),

                type:
                    (source.type || "").trim(),

                reference:
                    (source.reference || "").trim(),

                description:
                    (source.description || "").trim()
            })
        );
}


// =========================================================
// ANALIZAR DOCUMENTACIÓN
// =========================================================

async function analyzeDocuments(button = null) {
    const freeText =
        ($("freeText")?.value || "").trim();

    if (
        !state.files.length &&
        !freeText
    ) {
        notify(
            "Cargue al menos un archivo o texto libre.",
            "warning"
        );

        return;
    }

    if (button) {
        button.disabled = true;
        button.textContent =
            "Analizando archivos...";
    }

    try {
        const formData =
            new FormData();

        state.files.forEach((item) => {
            formData.append(
                "files",
                item.file,
                item.name
            );
        });

        formData.append(
            "freeText",
            freeText
        );

        const response = await fetch(
            "/analyze",
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
            const responseText =
                await response.text();

            console.error(
                "Respuesta inesperada del backend:",
                responseText
            );

            throw new Error(
                "El servidor devolvió una respuesta inválida. Revisá el deploy y los logs de Render."
            );
        }

        const data =
            await response.json();

        if (!response.ok) {
            throw new Error(
                data.error ||
                "No se pudo analizar la documentación."
            );
        }

        state.facts =
            data.facts || [];

        renderFacts();
        updateDashboard();

        if (
            Array.isArray(data.errors) &&
            data.errors.length
        ) {
            notify(
                (
                    data.message ||
                    "Documentación procesada."
                )
                +
                " Observaciones: "
                +
                data.errors.join(" | "),
                "warning"
            );
        } else {
            notify(
                data.message ||
                "Documentación analizada."
            );
        }

    } catch (error) {
        notify(
            error.message,
            "error"
        );

    } finally {
        if (button) {
            button.disabled = false;
            button.textContent =
                "🔍 Analizar documentación";
        }
    }
}


// =========================================================
// HECHOS
// =========================================================

function renderFacts() {
    const container =
        $("factsContainer");

    if (!container) {
        return;
    }

    if (!state.facts.length) {
        container.innerHTML =
            '<p class="section-description">No se identificaron elementos verificables.</p>';

        return;
    }

    container.innerHTML = state.facts
        .map((fact) => {
            const status =
                fact.status || "pending";

            const statusText =
                status === "accepted"
                    ? "Aceptado"
                    : status === "discarded"
                        ? "Descartado"
                        : "Pendiente";

            return `
            <article
                class="fact-card ${status}"
            >
                <div class="fact-title">
                    ${escapeHtml(fact.description)}

                    <span
                        class="status-badge status-${status}"
                    >
                        ${statusText}
                    </span>
                </div>

                <div class="fact-value">
                    ${escapeHtml(
                        fact.value ||
                        "Información no identificada."
                    )}
                </div>

                <div class="fact-meta">
                    <strong>
                        Fuente:
                    </strong>

                    ${escapeHtml(
                        fact.source || "N/A"
                    )}

                    |

                    <strong>
                        Referencia:
                    </strong>

                    ${escapeHtml(
                        fact.reference || "N/A"
                    )}
                </div>

                <div class="fact-actions">
                    <button
                        class="btn btn-success btn-sm"
                        type="button"
                        onclick="setFactStatus(${fact.id}, 'accepted')"
                    >
                        Aceptar
                    </button>

                    <button
                        class="btn btn-secondary btn-sm"
                        type="button"
                        onclick="editFact(${fact.id})"
                    >
                        Editar
                    </button>

                    <button
                        class="btn btn-danger btn-sm"
                        type="button"
                        onclick="setFactStatus(${fact.id}, 'discarded')"
                    >
                        Descartar
                    </button>
                </div>
            </article>
            `;
        })
        .join("");
}

function setFactStatus(id, status) {
    const fact = state.facts.find(
        (item) => item.id === id
    );

    if (fact) {
        fact.status = status;

        renderFacts();
        updateDashboard();
    }
}

function editFact(id) {
    const fact = state.facts.find(
        (item) => item.id === id
    );

    if (!fact) {
        return;
    }

    const description =
        prompt(
            "Descripción:",
            fact.description
        );

    if (description === null) {
        return;
    }

    const value =
        prompt(
            "Valor:",
            fact.value || ""
        );

    if (value === null) {
        return;
    }

    const source =
        prompt(
            "Fuente:",
            fact.source || ""
        );

    if (source === null) {
        return;
    }

    const reference =
        prompt(
            "Referencia:",
            fact.reference || ""
        );

    if (reference === null) {
        return;
    }

    Object.assign(
        fact,
        {
            description,
            value,
            source,
            reference,
            status: "accepted"
        }
    );

    renderFacts();
    updateDashboard();
}


// =========================================================
// TAREAS
// =========================================================

function addTask(data = {}) {
    state.tasks.push({
        id: generateId(),
        ...data
    });

    renderTasks();
    updateDashboard();
}

function renderTasks() {
    const container =
        $("tasksContainer");

    if (!container) {
        return;
    }

    if (!state.tasks.length) {
        container.innerHTML =
            '<p class="section-description">Todavía no agregaste tareas.</p>';

        return;
    }

    container.innerHTML = state.tasks
        .map(
            (
                task,
                index
            ) => `
            <div
                class="task-card"
                data-id="${task.id}"
            >
                <div class="task-card-header">
                    <h4>
                        Tarea ${String(index + 1).padStart(2, "0")}
                    </h4>

                    <button
                        class="btn btn-danger btn-sm"
                        type="button"
                        onclick="deleteTask('${task.id}')"
                    >
                        Eliminar
                    </button>
                </div>

                <div class="form-group">
                    <label>
                        Descripción de la tarea
                    </label>

                    <textarea
                        data-field="descripcion"
                        placeholder="Ej. Se solicitó al área el detalle de préstamos otorgados..."
                        oninput="syncTask('${task.id}', this)"
                    >${escapeHtml(task.descripcion || "")}</textarea>
                </div>

                <div class="inline-grid">
                    <div class="form-group">
                        <label>
                            Fuente
                        </label>

                        <input
                            data-field="fuente"
                            value="${escapeAttribute(task.fuente || "")}"
                            oninput="syncTask('${task.id}', this)"
                        >
                    </div>

                    <div class="form-group">
                        <label>
                            Evidencia
                        </label>

                        <input
                            data-field="evidencia"
                            value="${escapeAttribute(task.evidencia || "")}"
                            oninput="syncTask('${task.id}', this)"
                        >
                    </div>
                </div>

                <div class="inline-grid">
                    <div class="form-group">
                        <label>
                            Resultado de la prueba
                        </label>

                        <textarea
                            data-field="resultado"
                            oninput="syncTask('${task.id}', this)"
                        >${escapeHtml(task.resultado || "")}</textarea>
                    </div>

                    <div class="form-group">
                        <label>
                            Referencia
                        </label>

                        <input
                            data-field="referencia"
                            value="${escapeAttribute(task.referencia || "")}"
                            oninput="syncTask('${task.id}', this)"
                        >
                    </div>
                </div>
            </div>
            `
        )
        .join("");
}

function syncTask(id, element) {
    const task = state.tasks.find(
        (item) => item.id === id
    );

    if (task) {
        task[element.dataset.field] =
            element.value;
    }
}

function deleteTask(id) {
    state.tasks =
        state.tasks.filter(
            (task) => task.id !== id
        );

    renderTasks();
    updateDashboard();
}


// =========================================================
// RESULTADOS
// =========================================================

function addResult(data = {}) {
    state.results.push({
        id: generateId(),
        clasificacion: "Observación",
        ...data
    });

    renderResults();
}

function renderResults() {
    const container =
        $("resultsContainer");

    if (!container) {
        return;
    }

    if (!state.results.length) {
        container.innerHTML =
            '<p class="section-description">Todavía no agregaste resultados.</p>';

        return;
    }

    container.innerHTML = state.results
        .map(
            (
                result,
                index
            ) => `
            <div
                class="result-card"
                data-id="${result.id}"
            >
                <div class="result-card-header">
                    <h4>
                        Resultado ${String(index + 1).padStart(2, "0")}
                    </h4>

                    <button
                        class="btn btn-danger btn-sm"
                        type="button"
                        onclick="deleteResult('${result.id}')"
                    >
                        Eliminar
                    </button>
                </div>

                <div class="inline-grid">
                    <div class="form-group">
                        <label>
                            Concepto
                        </label>

                        <input
                            data-field="concepto"
                            value="${escapeAttribute(result.concepto || "")}"
                            oninput="syncResult('${result.id}', this)"
                        >
                    </div>

                    <div class="form-group">
                        <label>
                            Clasificación
                        </label>

                        <select
                            data-field="clasificacion"
                            onchange="syncResult('${result.id}', this)"
                        >
                            ${[
                                "Hallazgo",
                                "Observación",
                                "Oportunidad de mejora",
                                "Sin excepción",
                                "Acción ya implementada"
                            ]
                                .map(
                                    (value) => `
                                    <option
                                        value="${value}"
                                        ${result.clasificacion === value ? "selected" : ""}
                                    >
                                        ${value}
                                    </option>
                                    `
                                )
                                .join("")
                            }
                        </select>
                    </div>
                </div>

                <div class="inline-grid">
                    <div class="form-group">
                        <label>
                            Cantidad
                        </label>

                        <input
                            data-field="cantidad"
                            value="${escapeAttribute(result.cantidad || "")}"
                            oninput="syncResult('${result.id}', this)"
                        >
                    </div>

                    <div class="form-group">
                        <label>
                            Importe
                        </label>

                        <input
                            data-field="importe"
                            value="${escapeAttribute(result.importe || "")}"
                            oninput="syncResult('${result.id}', this)"
                        >
                    </div>

                    <div class="form-group">
                        <label>
                            Porcentaje
                        </label>

                        <input
                            data-field="porcentaje"
                            value="${escapeAttribute(result.porcentaje || "")}"
                            oninput="syncResult('${result.id}', this)"
                        >
                    </div>

                    <div class="form-group">
                        <label>
                            Observación
                        </label>

                        <input
                            data-field="observacion"
                            value="${escapeAttribute(result.observacion || "")}"
                            oninput="syncResult('${result.id}', this)"
                        >
                    </div>
                </div>
            </div>
            `
        )
        .join("");
}

function syncResult(id, element) {
    const result = state.results.find(
        (item) => item.id === id
    );

    if (result) {
        result[element.dataset.field] =
            element.value;
    }
}

function deleteResult(id) {
    state.results =
        state.results.filter(
            (result) => result.id !== id
        );

    renderResults();
}


// =========================================================
// HALLAZGOS
// =========================================================

function addFinding(data = {}) {
    state.findings.push({
        id: generateId(),
        criticidad: "medium",
        estado: "Pendiente",
        ...data
    });

    renderFindings();
    updateDashboard();
}

function renderFindings() {
    const container =
        $("findingsContainer");

    if (!container) {
        return;
    }

    if (!state.findings.length) {
        container.innerHTML =
            '<p class="section-description">Todavía no agregaste hallazgos.</p>';

        return;
    }

    container.innerHTML = state.findings
        .map(
            (
                finding,
                index
            ) => {
                const criticality =
                    finding.criticidad || "medium";

                return `
                <article
                    class="finding-card ${criticality}"
                    data-id="${finding.id}"
                >
                    <div class="finding-header">
                        <div>
                            <span class="finding-number">
                                Hallazgo ${String(index + 1).padStart(2, "0")}
                            </span>

                            <h4>
                                ${escapeHtml(
                                    finding.titulo ||
                                    "Nuevo hallazgo"
                                )}
                            </h4>
                        </div>

                        <div>
                            <span
                                class="criticality-badge criticality-${criticality}"
                            >
                                ${
                                    criticality === "high"
                                        ? "🔴 Alto"
                                        : criticality === "medium"
                                            ? "🟡 Medio"
                                            : "🟢 Bajo"
                                }
                            </span>

                            <button
                                class="btn btn-danger btn-sm"
                                type="button"
                                onclick="deleteFinding('${finding.id}')"
                            >
                                Eliminar
                            </button>
                        </div>
                    </div>

                    <div class="form-grid">
                        <div class="form-group full-width">
                            <label>
                                Título
                            </label>

                            <input
                                data-field="titulo"
                                value="${escapeAttribute(finding.titulo || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>

                        <div class="form-group full-width">
                            <label>
                                Descripción
                            </label>

                            <textarea
                                data-field="descripcion"
                                oninput="syncFinding('${finding.id}', this)"
                            >${escapeHtml(finding.descripcion || "")}</textarea>
                        </div>

                        <div class="form-group full-width">
                            <label>
                                Condición / hecho observado
                            </label>

                            <textarea
                                data-field="condicion"
                                oninput="syncFinding('${finding.id}', this)"
                            >${escapeHtml(finding.condicion || "")}</textarea>
                        </div>

                        <div class="form-group">
                            <label>
                                Área responsable *
                            </label>

                            <input
                                data-field="area_responsable"
                                list="areasResponsables"
                                value="${escapeAttribute(finding.area_responsable || "")}"
                                placeholder="Escriba o seleccione un área"
                                oninput="syncFinding('${finding.id}', this)"
                            >

                            <datalist id="areasResponsables">
                                <option value="Contabilidad">
                                <option value="Créditos">
                                <option value="Sistemas">
                                <option value="Compras">
                                <option value="RR.HH.">
                                <option value="Operaciones">
                                <option value="Tesorería">
                                <option value="Comercial">
                                <option value="Logística">
                            </datalist>
                        </div>

                        <div class="form-group">
                            <label>
                                Responsable del plan de acción
                            </label>

                            <input
                                data-field="responsable_plan"
                                value="${escapeAttribute(finding.responsable_plan || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>

                        <div class="form-group">
                            <label>
                                Criticidad *
                            </label>

                            <select
                                data-field="criticidad"
                                onchange="
                                    syncFinding('${finding.id}', this);
                                    renderFindings();
                                    updateDashboard();
                                "
                            >
                                <option
                                    value="high"
                                    ${criticality === "high" ? "selected" : ""}
                                >
                                    Alto
                                </option>

                                <option
                                    value="medium"
                                    ${criticality === "medium" ? "selected" : ""}
                                >
                                    Medio
                                </option>

                                <option
                                    value="low"
                                    ${criticality === "low" ? "selected" : ""}
                                >
                                    Bajo
                                </option>
                            </select>
                        </div>

                        <div class="form-group">
                            <label>
                                Estado
                            </label>

                            <select
                                data-field="estado"
                                onchange="syncFinding('${finding.id}', this)"
                            >
                                ${[
                                    "Pendiente",
                                    "En análisis",
                                    "En curso",
                                    "Implementado",
                                    "Cerrado"
                                ]
                                    .map(
                                        (value) => `
                                        <option
                                            value="${value}"
                                            ${finding.estado === value ? "selected" : ""}
                                        >
                                            ${value}
                                        </option>
                                        `
                                    )
                                    .join("")
                                }
                            </select>
                        </div>

                        <div class="form-group">
                            <label>
                                Fecha objetivo
                            </label>

                            <input
                                type="date"
                                data-field="fecha_objetivo"
                                value="${escapeAttribute(finding.fecha_objetivo || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>

                        <div class="form-group">
                            <label>
                                Fuente
                            </label>

                            <input
                                data-field="fuente"
                                value="${escapeAttribute(finding.fuente || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>

                        <div class="form-group full-width">
                            <label>
                                Riesgo *
                            </label>

                            <textarea
                                data-field="riesgo"
                                oninput="syncFinding('${finding.id}', this)"
                            >${escapeHtml(finding.riesgo || "")}</textarea>
                        </div>

                        <div class="form-group full-width">
                            <label>
                                Propuesta de mejora *
                            </label>

                            <textarea
                                data-field="propuesta_mejora"
                                oninput="syncFinding('${finding.id}', this)"
                            >${escapeHtml(finding.propuesta_mejora || "")}</textarea>
                        </div>

                        <div class="form-group">
                            <label>
                                Fundamento cuantitativo
                            </label>

                            <input
                                data-field="fundamento_cuantitativo"
                                value="${escapeAttribute(finding.fundamento_cuantitativo || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>

                        <div class="form-group">
                            <label>
                                Evidencia
                            </label>

                            <input
                                data-field="evidencia"
                                value="${escapeAttribute(finding.evidencia || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>

                        <div class="form-group">
                            <label>
                                Referencia / Ticket
                            </label>

                            <input
                                data-field="referencia"
                                value="${escapeAttribute(finding.referencia || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>

                        <div class="form-group">
                            <label>
                                Seguimiento
                            </label>

                            <input
                                data-field="seguimiento"
                                value="${escapeAttribute(finding.seguimiento || "")}"
                                oninput="syncFinding('${finding.id}', this)"
                            >
                        </div>
                    </div>
                </article>
                `;
            }
        )
        .join("");
}

function syncFinding(id, element) {
    const finding = state.findings.find(
        (item) => item.id === id
    );

    if (finding) {
        finding[element.dataset.field] =
            element.value;
    }
}

function deleteFinding(id) {
    state.findings =
        state.findings.filter(
            (finding) => finding.id !== id
        );

    renderFindings();
    updateDashboard();
}


// =========================================================
// DATOS DEL TRABAJO
// =========================================================

function collectAuditData() {
    return {
        titulo:
            ($("auditTitle")?.value || "").trim(),

        analisis:
            ($("analysis")?.value || "").trim(),

        sector:
            ($("sector")?.value || "").trim(),

        proceso:
            ($("process")?.value || "").trim(),

        periodo:
            ($("period")?.value || "").trim(),

        alcance:
            ($("scope")?.value || "").trim(),

        fecha:
            $("workDate")?.value || "",

        auditor:
            ($("auditor")?.value || "").trim(),

        contexto:
            ($("context")?.value || "").trim(),

        instrucciones:
            ($("writingInstructions")?.value || "").trim(),

        objetivos:
            getObjectives(),

        fuentes:
            getSources()
    };
}


// =========================================================
// VALIDACIONES
// =========================================================

function validateBeforeMemo() {
    const audit =
        collectAuditData();

    const messages = [];

    if (!audit.objetivos.length) {
        messages.push(
            "Falta al menos un objetivo."
        );
    }

    if (!audit.periodo) {
        messages.push(
            "Falta el período."
        );
    }

    if (
        !state.facts.some(
            (fact) =>
                fact.status === "accepted"
        )
    ) {
        messages.push(
            "Debe aceptar al menos un hecho."
        );
    }

    state.findings.forEach(
        (
            finding,
            index
        ) => {
            const prefix =
                `Hallazgo ${index + 1}: `;

            if (!finding.area_responsable) {
                messages.push(
                    `${prefix}falta Área responsable.`
                );
            }

            if (!finding.riesgo) {
                messages.push(
                    `${prefix}falta Riesgo.`
                );
            }

            if (!finding.propuesta_mejora) {
                messages.push(
                    `${prefix}falta Propuesta de mejora.`
                );
            }
        }
    );

    if (messages.length) {
        notify(
            messages.join(" "),
            "warning"
        );

        return false;
    }

    return true;
}


// =========================================================
// GENERAR MEMO
// =========================================================

async function generateMemo() {
    if (!validateBeforeMemo()) {
        return;
    }

    const acceptedFacts =
        state.facts.filter(
            (fact) =>
                fact.status === "accepted"
        );

    try {
        const response =
            await fetch(
                "/generate-memo",
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({
                            auditData:
                                collectAuditData(),

                            validatedFacts:
                                acceptedFacts,

                            tasks:
                                state.tasks,

                            results:
                                state.results,

                            findings:
                                state.findings,

                            style:
                                $("memoStyle")?.value ||
                                "ejecutivo"
                        })
                }
            );

        const data =
            await response.json();

        if (!response.ok) {
            throw new Error(
                data.error ||
                "No se pudo generar el memo."
            );
        }

        state.memo =
            data.memo;

        renderMemo(
            state.memo
        );

        notify(
            data.message ||
            "Memo generado."
        );

    } catch (error) {
        notify(
            error.message,
            "error"
        );
    }
}


// =========================================================
// PREVIEW DEL MEMO
// =========================================================

function renderMemo(memo) {
    const header =
        memo.header || {};

    const criticalityLabel = {
        high: "Alto",
        medium: "Medio",
        low: "Bajo"
    };

    const objectives =
        (memo.objetivos || [])
            .map(
                (objective) =>
                    `<li>${escapeHtml(objective)}</li>`
            )
            .join("")
        ||
        "<li>Sin objetivos.</li>";

    const tasks =
        (memo.tareas || [])
            .map(
                (task) => `
                <li>
                    <strong>
                        ${escapeHtml(task.descripcion || "Tarea")}
                    </strong>

                    ${
                        task.resultado
                            ? ": " +
                              escapeHtml(task.resultado)
                            : ""
                    }
                </li>
                `
            )
            .join("")
        ||
        "<li>Sin tareas.</li>";

    const facts =
        (memo.hechos_validados || [])
            .map(
                (fact) => `
                <li>
                    ${escapeHtml(fact.description)}:

                    <strong>
                        ${escapeHtml(fact.value)}
                    </strong>

                    — Fuente:
                    ${escapeHtml(fact.source)}

                    ${
                        fact.reference
                            ? " — Ref.: " +
                              escapeHtml(fact.reference)
                            : ""
                    }
                </li>
                `
            )
            .join("");

    const resultRows =
        (memo.resultados || [])
            .map(
                (result) => `
                <tr>
                    <td>
                        ${escapeHtml(result.concepto || "")}
                    </td>

                    <td>
                        ${escapeHtml(result.cantidad || "")}
                    </td>

                    <td>
                        ${escapeHtml(result.importe || "")}
                    </td>

                    <td>
                        ${escapeHtml(result.porcentaje || "")}
                    </td>

                    <td>
                        ${escapeHtml(result.clasificacion || "")}
                    </td>

                    <td>
                        ${escapeHtml(result.observacion || "")}
                    </td>
                </tr>
                `
            )
            .join("");

    const findings =
        (memo.hallazgos || [])
            .map(
                (
                    finding,
                    index
                ) => `
                <div
                    class="memo-finding ${finding.criticidad || "medium"}"
                >
                    <strong>
                        Hallazgo ${String(index + 1).padStart(2, "0")}
                        –
                        ${escapeHtml(
                            finding.titulo ||
                            "Sin título"
                        )}
                    </strong>

                    <br>

                    <span
                        class="criticality-badge criticality-${finding.criticidad || "medium"}"
                    >
                        ${criticalityLabel[finding.criticidad] || "N/A"}
                    </span>

                    <p>
                        <strong>Descripción:</strong>
                        ${nl2br(finding.descripcion || "N/A")}
                    </p>

                    <p>
                        <strong>Condición:</strong>
                        ${nl2br(finding.condicion || "N/A")}
                    </p>

                    <p>
                        <strong>Riesgo:</strong>
                        ${nl2br(finding.riesgo || "N/A")}
                    </p>

                    <p>
                        <strong>Área responsable:</strong>
                        ${escapeHtml(finding.area_responsable || "N/A")}
                    </p>

                    <p>
                        <strong>Responsable plan:</strong>
                        ${escapeHtml(finding.responsable_plan || "N/A")}
                    </p>

                    <p>
                        <strong>Propuesta de mejora:</strong>
                        ${nl2br(finding.propuesta_mejora || "N/A")}
                    </p>

                    <p>
                        <strong>Estado:</strong>
                        ${escapeHtml(finding.estado || "Pendiente")}
                    </p>

                    <p>
                        <strong>Fecha objetivo:</strong>
                        ${escapeHtml(finding.fecha_objetivo || "N/A")}
                    </p>

                    <p>
                        <strong>Fundamento:</strong>
                        ${escapeHtml(finding.fundamento_cuantitativo || "N/A")}
                    </p>

                    <p>
                        <strong>Fuente / evidencia:</strong>
                        ${escapeHtml(finding.fuente || "")}
                        ${escapeHtml(finding.evidencia || "")}
                    </p>
                </div>
                `
            )
            .join("")
        ||
        "<p>Sin hallazgos.</p>";

    const memoPreview =
        $("memoPreview");

    if (!memoPreview) {
        return;
    }

    memoPreview.innerHTML = `
        <div
            class="memo-document"
            contenteditable="true"
        >
            <h2>
                MEMO DE AUDITORÍA INTERNA
            </h2>

            <table class="memo-header-table">
                <tr>
                    <td>Título</td>
                    <td>${escapeHtml(header.titulo || "")}</td>
                </tr>

                <tr>
                    <td>Análisis</td>
                    <td>${escapeHtml(header.analisis || "")}</td>
                </tr>

                <tr>
                    <td>Sector</td>
                    <td>${escapeHtml(header.sector || "")}</td>
                </tr>

                <tr>
                    <td>Proceso</td>
                    <td>${escapeHtml(header.proceso || "")}</td>
                </tr>

                <tr>
                    <td>Período</td>
                    <td>${escapeHtml(header.periodo || "")}</td>
                </tr>

                <tr>
                    <td>Alcance</td>
                    <td>${escapeHtml(header.alcance || "")}</td>
                </tr>

                <tr>
                    <td>Auditor</td>
                    <td>${escapeHtml(header.auditor || "")}</td>
                </tr>

                <tr>
                    <td>Fecha</td>
                    <td>${escapeHtml(header.fecha || "")}</td>
                </tr>
            </table>

            <h3>
                1. Objetivo
            </h3>

            <ol>
                ${objectives}
            </ol>

            <h3>
                2. Alcance
            </h3>

            <p>
                ${escapeHtml(header.alcance || "N/A")}
            </p>

            <h3>
                3. Tareas realizadas
            </h3>

            <ol>
                ${tasks}
            </ol>

            <h3>
                4. Resultados / hechos validados
            </h3>

            <ul>
                ${facts}
            </ul>

            ${
                resultRows
                    ? `
                    <table class="memo-result-table">
                        <thead>
                            <tr>
                                <th>Concepto</th>
                                <th>Cantidad</th>
                                <th>Importe</th>
                                <th>%</th>
                                <th>Clasificación</th>
                                <th>Observación</th>
                            </tr>
                        </thead>

                        <tbody>
                            ${resultRows}
                        </tbody>
                    </table>
                    `
                    : ""
            }

            <h3>
                5. Hallazgos
            </h3>

            ${findings}

            <h3>
                6. Conclusiones
            </h3>

            <p>
                ${nl2br(
                    memo.conclusiones ||
                    "N/A"
                )}
            </p>
        </div>
    `;
}


// =========================================================
// MEJORAR REDACCIÓN
// =========================================================

async function improveText() {
    const documentElement =
        document.querySelector(
            "#memoPreview .memo-document"
        );

    if (!documentElement) {
        notify(
            "Primero genere el memo.",
            "warning"
        );

        return;
    }

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

                    body:
                        JSON.stringify({
                            text:
                                documentElement.innerText,

                            style:
                                $("memoStyle")?.value ||
                                "ejecutivo"
                        })
                }
            );

        const data =
            await response.json();

        if (!response.ok) {
            throw new Error(
                data.error ||
                "No se pudo mejorar el texto."
            );
        }

        const originalText =
            documentElement.innerText;

        if (
            data.improved &&
            data.improved !== originalText
        ) {
            documentElement.innerText =
                data.improved;
        }

        notify(
            data.message ||
            "Redacción revisada.",
            data.improved === originalText
                ? "warning"
                : "success"
        );

    } catch (error) {
        notify(
            error.message,
            "error"
        );
    }
}


// =========================================================
// EXPORTAR
// =========================================================

async function exportToExcel() {
    if (!state.memo) {
        notify(
            "Primero genere el memo.",
            "warning"
        );

        return;
    }

    try {
        const response =
            await fetch(
                "/export-excel",
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({
                            memo:
                                state.memo
                        })
                }
            );

        if (!response.ok) {
            const data =
                await response
                    .json()
                    .catch(
                        () => ({})
                    );

            throw new Error(
                data.error ||
                "No se pudo exportar el Excel."
            );
        }

        const blob =
            await response.blob();

        const url =
            URL.createObjectURL(blob);

        const link =
            document.createElement("a");

        link.href =
            url;

        link.download =
            "audit_memo.xlsx";

        document.body.appendChild(
            link
        );

        link.click();
        link.remove();

        URL.revokeObjectURL(
            url
        );

        notify(
            "Excel exportado."
        );

    } catch (error) {
        notify(
            error.message,
            "error"
        );
    }
}


// =========================================================
// GUARDAR PROGRESO LOCAL
// =========================================================

function saveProgress() {
    const data = {
        files:
            state.files.map(
                (file) => ({
                    name: file.name,
                    type: file.type,
                    size: file.size
                })
            ),

        facts:
            state.facts,

        tasks:
            state.tasks,

        results:
            state.results,

        findings:
            state.findings,

        sources:
            state.sources,

        auditData:
            collectAuditData()
    };

    localStorage.setItem(
        "auditMemoProgress",
        JSON.stringify(data)
    );

    notify(
        "Progreso guardado en este navegador."
    );
}


// =========================================================
// CARGAR PROGRESO
// =========================================================

function loadProgress() {
    const saved =
        localStorage.getItem(
            "auditMemoProgress"
        );

    if (!saved) {
        return;
    }

    try {
        const data =
            JSON.parse(saved);

        /*
        Los archivos físicos no se pueden recuperar
        desde localStorage.
        */

        state.files = [];

        state.facts =
            data.facts || [];

        state.tasks =
            data.tasks || [];

        state.results =
            data.results || [];

        state.findings =
            data.findings || [];

        state.sources =
            data.sources || [];

        if (data.auditData) {
            const audit =
                data.auditData;

            if ($("auditTitle")) {
                $("auditTitle").value =
                    audit.titulo || "";
            }

            if ($("analysis")) {
                $("analysis").value =
                    audit.analisis || "";
            }

            if ($("sector")) {
                $("sector").value =
                    audit.sector || "";
            }

            if ($("process")) {
                $("process").value =
                    audit.proceso || "";
            }

            if ($("period")) {
                $("period").value =
                    audit.periodo || "";
            }

            if ($("scope")) {
                $("scope").value =
                    audit.alcance || "";
            }

            if ($("workDate")) {
                $("workDate").value =
                    audit.fecha || "";
            }

            if ($("auditor")) {
                $("auditor").value =
                    audit.auditor || "";
            }

            if ($("context")) {
                $("context").value =
                    audit.contexto || "";
            }

            if ($("writingInstructions")) {
                $("writingInstructions").value =
                    audit.instrucciones || "";
            }

            document
                .querySelectorAll(
                    ".objective-row"
                )
                .forEach(
                    (element) =>
                        element.remove()
                );

            (audit.objetivos || [])
                .forEach(
                    (objective) =>
                        addObjective(
                            objective
                        )
                );
        }

        renderFiles();
        renderSources();
        renderFacts();
        renderTasks();
        renderResults();
        renderFindings();
        updateDashboard();

        notify(
            "Progreso cargado. Para volver a analizar deberá seleccionar nuevamente los archivos originales.",
            "warning"
        );

    } catch (error) {
        console.error(
            "Error al recuperar progreso:",
            error
        );
    }
}


// =========================================================
// NUEVO MEMO
// =========================================================

function newMemo() {
    const confirmReset =
        confirm(
            "¿Querés iniciar un nuevo memo? Se eliminará el progreso guardado en este navegador."
        );

    if (!confirmReset) {
        return;
    }

    localStorage.removeItem(
        "auditMemoProgress"
    );

    location.reload();
}


// =========================================================
// INICIO
// =========================================================

document.addEventListener(
    "DOMContentLoaded",
    () => {
        initializeUpload();

        if ($("workDate")) {
            $("workDate").value =
                new Date()
                    .toISOString()
                    .slice(0, 10);
        }

        addObjective();

        renderSources();
        renderFiles();
        renderFacts();
        renderTasks();
        renderResults();
        renderFindings();
        updateDashboard();

        loadProgress();
    }
);
