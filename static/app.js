// ============================================================
// AUDIT MEMO BUILDER - FRONTEND ESTABLE
// ============================================================

const STORAGE_PREFIX = "auditMemoBuilder:v3";
const AUTOSAVE_DELAY = 350;
let autosaveTimer = null;
let selectedFiles = [];
let extractionInProgress = false;

function createEmptyState() {
    return {
        auditId: crypto.randomUUID(),
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        currentStep: 1,
        general: {
            title: "", area: "", process: "", period: "",
            auditor: "", objective: "", scope: ""
        },
        sources: [],
        extracted: [],
        findings: []
    };
}

let state = createEmptyState();

const GENERAL_FIELDS = {
    auditTitle: "title",
    auditArea: "area",
    auditProcess: "process",
    auditPeriod: "period",
    auditAuditor: "auditor",
    auditObjective: "objective",
    auditScope: "scope"
};

function el(id) { return document.getElementById(id); }
function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function storageKey() { return `${STORAGE_PREFIX}:current`; }
function onlyHallazgos(items) {
    return Array.isArray(items) ? items : [];
}
function indexHallazgos(items) {
    return (Array.isArray(items) ? items : [])
        .map((item, index) => ({ item, index }));
}
function hallazgoItemsWithIndexes() {
    return indexHallazgos(state.extracted);
}
function isFindingEligible(item) {
    return Boolean(item && item.category);
}
function eligibleIncludedHallazgos(items) {
    return (Array.isArray(items) ? items : []).filter(item => item.included && !item.converted);
}
// Solo los marcados EXPLÍCITAMENTE como hallazgo por el auditor
function eligibleSelectedFindings(items) {
    return (Array.isArray(items) ? items : []).filter(item => item.selectedAsFinding && !item.converted);
}
// Busca el candidato original de un hallazgo por sourceItemId
function getSourceItemForFinding(finding) {
    return (typeof state !== "undefined" ? state.extracted : []).find(item => item.id === finding.sourceItemId) || null;
}
// Reconcilia una nueva extracción con los candidatos anteriores para preservar IDs
function reconcileExtracted(incoming, previous) {
    const prevMap = new Map();
    (previous || []).forEach(item => {
        const key = `${item.filename || ""}|${item.originName || ""}|${item.reference || ""}`;
        if (!prevMap.has(key)) prevMap.set(key, item);
    });
    return incoming.map(item => {
        const key = `${item.filename || ""}|${item.originName || ""}|${item.reference || ""}`;
        const prev = prevMap.get(key);
        if (prev) {
            // Reutilizar el id anterior para no romper sourceItemId en hallazgos ya creados
            return {
                ...item,
                id: prev.id,
                included: prev.included,
                selectedAsFinding: prev.selectedAsFinding || false,
                converted: prev.converted,
            };
        }
        return { ...item, selectedAsFinding: Boolean(item.selectedAsFinding) };
    });
}

function saveState() {
    try {
        state.updatedAt = new Date().toISOString();
        localStorage.setItem(storageKey(), JSON.stringify(state));
    } catch (error) {
        console.warn("No se pudo guardar el memo", error);
    }
}
function scheduleSave() {
    clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(saveState, AUTOSAVE_DELAY);
}
function loadState() {
    try {
        const raw = localStorage.getItem(storageKey());
        if (!raw) return;
        const saved = JSON.parse(raw);
        if (!saved || typeof saved !== "object") return;
        state = {
            ...createEmptyState(),
            ...saved,
            general: { ...createEmptyState().general, ...(saved.general || {}) },
            sources: saved.sources || [],
            extracted: onlyHallazgos(saved.extracted),
            findings: saved.findings || []
        };
    } catch (error) {
        console.warn("No se pudo recuperar el memo", error);
    }
}

function startNewAudit(keepAuditor = true) {
    const auditor = keepAuditor ? state.general.auditor : "";
    state = createEmptyState();
    state.general.auditor = auditor;
    selectedFiles = [];
    saveState();
    hydrateGeneral();
    renderFiles();
    renderExtraction();
    renderFindings();
    renderMemoPreview();
    renderValidation();
    goToStep(1);
    showToast("Se inició una auditoría nueva.", "success");
}

function hydrateGeneral() {
    Object.entries(GENERAL_FIELDS).forEach(([id, key]) => {
        if (el(id)) el(id).value = state.general[key] || "";
    });
}

function bindGeneralFields() {
    Object.entries(GENERAL_FIELDS).forEach(([id, key]) => {
        const input = el(id);
        if (!input) return;
        input.addEventListener("input", () => {
            state.general[key] = input.value || "";
            scheduleSave();
            renderMemoPreview();
            renderValidation();
        });
    });
}

// ============================================================
// NAVEGACIÓN - SIEMPRE LIBRE, INCLUSO DURANTE EXTRACCIÓN
// ============================================================

function goToStep(step) {
    const target = Math.max(1, Math.min(5, Number(step) || 1));
    state.currentStep = target;

    document.querySelectorAll(".page-step").forEach(section => {
        section.classList.toggle("active", section.id === `step${target}`);
    });
    document.querySelectorAll(".step-item").forEach(button => {
        const number = Number(button.dataset.step);
        button.classList.toggle("active", number === target);
        button.classList.toggle("completed", number < target);
    });

    if (target === 2) renderExtraction();
    if (target === 3) renderFindings();
    if (target === 4) renderMemoPreview();
    if (target === 5) renderValidation();
    scheduleSave();
    window.scrollTo({ top: 0, behavior: "smooth" });
}

// ============================================================
// ARCHIVOS
// ============================================================

function detectSourceType(name) {
    const ext = (name.split(".").pop() || "").toLowerCase();
    return ({xlsx: "Excel", csv: "CSV", docx: "Word", pdf: "PDF", txt: "TXT"})[ext] || ext.toUpperCase();
}
function fileSignature(file) { return `${file.name}::${file.size}::${file.lastModified}`; }
function formatFileSize(bytes) {
    const mb = bytes / (1024 * 1024);
    return mb >= 1 ? `${mb.toFixed(1)} MB` : `${Math.round(bytes / 1024)} KB`;
}

function addFiles(fileList) {
    const incoming = Array.from(fileList || []);
    if (!incoming.length) return;

    if ((state.extracted.length || state.findings.length) && selectedFiles.length === 0) {
        const newAudit = window.confirm(
            "Este memo ya contiene información.\n\n" +
            "¿Los archivos que estás cargando corresponden a una auditoría NUEVA?\n\n" +
            "Aceptar = comenzar limpio.\nCancelar = agregarlos al memo actual."
        );
        if (newAudit) startNewAudit(true);
    }

    const signatures = new Set(selectedFiles.map(fileSignature));
    incoming.forEach(file => {
        const signature = fileSignature(file);
        if (!signatures.has(signature)) {
            selectedFiles.push(file);
            signatures.add(signature);
        }
    });

    state.sources = selectedFiles.map(file => ({
        name: file.name,
        type: detectSourceType(file.name),
        reference: "",
        description: "Papel de trabajo aportado para la auditoría."
    }));
    saveState();
    renderFiles();
    renderValidation();
}

function removeFile(index) {
    selectedFiles.splice(index, 1);
    state.sources = selectedFiles.map(file => ({
        name: file.name,
        type: detectSourceType(file.name),
        reference: "",
        description: "Papel de trabajo aportado para la auditoría."
    }));
    saveState();
    renderFiles();
}

function renderFiles() {
    const container = el("fileList");
    if (!container) return;
    if (!selectedFiles.length) {
        container.innerHTML = state.sources.length
            ? `<div class="file-item"><div class="file-info"><strong>${state.sources.length} fuente(s) guardada(s)</strong><span>Volvé a seleccionar los archivos solo si querés ejecutar una nueva extracción.</span></div></div>`
            : "";
        return;
    }
    container.innerHTML = selectedFiles.map((file, index) => `
        <div class="file-item">
            <div class="file-info"><strong>${escapeHtml(file.name)}</strong><span>${formatFileSize(file.size)}</span></div>
            <button type="button" class="icon-button" onclick="removeFile(${index})">×</button>
        </div>
    `).join("");
}

function setupDropZone() {
    const input = el("fileInput");
    const zone = el("dropZone");
    if (input) {
        input.addEventListener("change", event => {
            addFiles(event.target.files);
            input.value = "";
        });
    }
    if (!zone) return;
    ["dragenter", "dragover"].forEach(name => zone.addEventListener(name, event => {
        event.preventDefault();
        zone.classList.add("dragging");
    }));
    ["dragleave", "drop"].forEach(name => zone.addEventListener(name, event => {
        event.preventDefault();
        zone.classList.remove("dragging");
    }));
    zone.addEventListener("drop", event => addFiles(event.dataTransfer.files));
}

// ============================================================
// EXTRACCIÓN
// ============================================================

function extractionButton() {
    return document.querySelector('button[onclick*="extractInformation"]');
}

function renderExtractionStatus(message, tone = "info") {
    const empty = el("extractionEmpty");
    if (!empty) return;
    empty.style.display = "block";
    empty.innerHTML = `
        <div class="empty-icon">${tone === "loading" ? "…" : "⌕"}</div>
        <h3>${escapeHtml(message)}</h3>
        ${tone === "loading" ? "<p>Podés moverte por las demás solapas mientras termina el análisis.</p>" : ""}
    `;
}

async function extractInformation() {
    const freeText = (el("freeText")?.value || "").trim();
    if (!selectedFiles.length && !freeText) {
        showToast("Cargá al menos un papel de trabajo o ingresá texto adicional.", "warning");
        return;
    }
    if (extractionInProgress) {
        showToast("La extracción ya está en curso.", "warning");
        return;
    }

    extractionInProgress = true;
    const button = extractionButton();
    if (button) {
        button.disabled = true;
        button.dataset.originalText = button.innerHTML;
        button.innerHTML = "Analizando archivo completo…";
    }

    goToStep(2);
    renderExtractionStatus("Iniciando análisis…", "loading");
    if (el("extractionList")) el("extractionList").innerHTML = "";
    if (el("extractedCount")) el("extractedCount").textContent = "Procesando…";

    const previousExtracted = [...state.extracted];
    state.extracted = [];

    const form = new FormData();
    selectedFiles.forEach(file => form.append("files", file));
    form.append("freeText", freeText);

    let renderPending = false;
    function scheduleRender() {
        if (renderPending) return;
        renderPending = true;
        requestAnimationFrame(() => {
            renderPending = false;
            renderExtraction();
            if (el("extractedCount")) {
                el("extractedCount").textContent = `${state.extracted.length} hallazgo(s)`;
            }
        });
    }

    try {
        const response = await fetch("/extract", { method: "POST", body: form });

        if (!response.ok) {
            let data = {};
            try { data = await response.json(); } catch (_) {}
            throw new Error(data.error || "No se pudo procesar la documentación.");
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";
        let warnings = [];

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split("\n");
            buffer = lines.pop();

            for (const line of lines) {
                if (!line.startsWith("data: ")) continue;
                try {
                    const event = JSON.parse(line.slice(6));
                    handleSSEEvent(event, previousExtracted, scheduleRender, () => { warnings = event.warnings || []; });
                } catch (parseErr) {}
            }
        }
        if (selectedFiles.length) {
            state.sources = selectedFiles.map(file => ({
                name: file.name,
                type: detectSourceType(file.name),
                reference: "",
                description: "Papel de trabajo aportado para la auditoría."
            }));
        }
        saveState();
        renderExtraction();
        renderValidation();
        if (el("extractedCount")) el("extractedCount").textContent = `${state.extracted.length} hallazgo(s)`;
    } catch (error) {
        console.error(error);
        renderExtractionStatus("La extracción no pudo completarse.");
        showToast(error.message || "Error durante la extracción.", "error");
    } finally {
        extractionInProgress = false;
        if (button) {
            button.disabled = false;
            if (button.dataset.originalText) button.innerHTML = button.dataset.originalText;
        }
    }
}

function handleSSEEvent(event, previousExtracted, scheduleRender, setWarnings) {
    if (event.type === "status") {
        renderExtractionStatus(event.message, "loading");
    } else if (event.type === "item") {
        const item = {
            ...event.item,
            included: Boolean(event.item.included),
            selectedAsFinding: Boolean(event.item.selectedAsFinding),
            converted: Boolean(event.item.converted)
        };
        const key = `${item.filename || ""}|${item.originName || ""}|${item.reference || ""}`;
        const prev = previousExtracted.find(p =>
            `${p.filename || ""}|${p.originName || ""}|${p.reference || ""}` === key
        );
        if (prev) {
            item.id = prev.id;
            item.included = prev.included;
            item.selectedAsFinding = prev.selectedAsFinding || false;
            item.converted = prev.converted;
        }
        state.extracted.push(item);
        scheduleRender();
    } else if (event.type === "error") {
        showToast(event.error, "warning");
    } else if (event.type === "done") {
        setWarnings();
        const warningText = (event.warnings || []).length ? ` · ${(event.warnings || []).length} advertencia(s)` : "";
        const hasErrors = (event.errors || []).length > 0;
        showToast(
            `${event.message || "Extracción finalizada"}${warningText}`,
            hasErrors ? "warning" : "success"
        );
    }
}

function renderExtraction() {
    const empty = el("extractionEmpty");
    const list = el("extractionList");
    const count = el("extractedCount");
    if (!empty || !list) return;

    if (extractionInProgress) {
        renderExtractionStatus("Analizando todas las solapas del papel de trabajo…", "loading");
        return;
    }
    const hallazgos = hallazgoItemsWithIndexes();
    if (count) count.textContent = `${hallazgos.length} hallazgo(s)`;
    if (!hallazgos.length) {
        empty.style.display = "block";
        empty.innerHTML = `<div class="empty-icon">⌕</div><h3>No se identificaron potenciales hallazgos</h3><p>La aplicación analiza todas las solapas en busca de inconsistencias o diferencias documentadas.</p>`;
        list.innerHTML = "";
        return;
    }


    empty.style.display = "none";
    list.innerHTML = hallazgos.map(({ item, index }) => {
        const eligible = isFindingEligible(item);
        const displayTitle = item.title || item.category || "Hallazgo detectado";
        const displaySituation = item.situation || item.text || "";
        const statusText = item.converted
            ? "✓ Convertido en hallazgo"
            : item.selectedAsFinding
                ? "★ Marcado como hallazgo (pendiente de convertir)"
                : eligible
                    ? "Puede convertirse en hallazgo"
                    : "Se incorporará como soporte del memo";
        return `
            <article class="extraction-card">
                <div class="extraction-card-header">
                    <div><span class="category-badge">${escapeHtml(item.category)}</span><strong class="source-title">${escapeHtml(item.filename)}</strong></div>
                    <div style="display:flex;gap:12px;align-items:center;">
                        <label class="include-check"><input type="checkbox" ${item.included ? "checked" : ""} onchange="toggleExtraction(${index}, this.checked)"> Incluir</label>
                        ${eligible ? `<label class="include-check" style="font-weight:600;color:#1e40af;"><input type="checkbox" ${item.selectedAsFinding ? "checked" : ""} onchange="toggleSelectedAsFinding('${item.id}', this.checked)"> Hallazgo</label>` : ""}
                    </div>
                </div>
                
                <div class="extraction-draft-box" style="margin: 10px 0; background: #f8fafc; padding: 12px; border-radius: 6px; border: 1px solid #e2e8f0;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <input
                            value="${escapeHtml(displayTitle)}"
                            oninput="updateExtractionField(${index}, 'title', this.value)"
                            style="width: 100%; font-weight: 600; margin-right: 12px;"
                        >
                        <button type="button" class="ai-button" onclick="draftExtractionItemWithAI(${index})">✦ Redactar con IA</button>
                    </div>
                    <div class="field" style="margin-bottom: 8px;">
                        <label style="font-size: 11px; font-weight: 600; color: #64748b; text-transform: uppercase;">Situación Observada (Redacción)</label>
                        <textarea class="extraction-text" style="width: 100%; min-height: 54px;" oninput="updateExtractionField(${index}, 'situation', this.value)">${escapeHtml(displaySituation)}</textarea>
                    </div>
                    <div class="form-grid">
                        <div class="field">
                            <label>Riesgo</label>
                            <textarea rows="3" oninput="updateExtractionField(${index}, 'risk', this.value)">${escapeHtml(item.risk || "")}</textarea>
                        </div>
                        <div class="field">
                            <label>Propuesta de mejora</label>
                            <textarea rows="3" oninput="updateExtractionField(${index}, 'proposal', this.value)">${escapeHtml(item.proposal || "")}</textarea>
                        </div>
                    </div>
                </div>

                <div class="trace-meta">
                    ${item.originName ? `<span>Solapa: <strong>${escapeHtml(item.originName)}</strong></span>` : ""}
                    ${item.reference ? `<span>${escapeHtml(item.reference)}</span>` : ""}
                    ${item.keyword ? `<span>Detectado por: ${escapeHtml(item.keyword)}</span>` : ""}
                </div>
                <div class="extraction-card-footer">
                    <span>${statusText}</span>
                    ${eligible && !item.converted ? `<button type="button" class="btn btn-secondary" onclick="convertOneToFinding('${item.id}')">Crear hallazgo</button>` : ""}
                </div>
            </article>`;
    }).join("");
}

function toggleExtraction(index, included) {
    if (!state.extracted[index]) return;
    state.extracted[index].included = included;
    scheduleSave();
    renderMemoPreview();
    renderValidation();
}

// Marca/desmarca un candidato como hallazgo. Busca por ID, nunca por índice.
function toggleSelectedAsFinding(itemId, value) {
    const item = state.extracted.find(i => i.id === itemId);
    if (!item) return;
    item.selectedAsFinding = value;
    scheduleSave();
    renderExtraction();
    renderValidation();
}

function updateExtractionText(index, value) {
    updateExtractionField(index, 'situation', value);
}

function updateExtractionField(index, field, value) {
    if (!state.extracted[index]) return;
    state.extracted[index][field] = value;
    if (field === 'situation') {
        state.extracted[index].text = value;
    }
    scheduleSave();
    renderMemoPreview();
}

async function draftExtractionItemWithAI(index) {
    const item = state.extracted[index];
    if (!item) return;

    showToast("Redactando hallazgo con IA…", "info");
    try {
        const response = await fetch("/draft-finding", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                text: item.situation || item.text,
                category: item.category,
                reason: item.reason || ""
            })
        });
        if (response.ok) {
            const data = await response.json();
            item.title = data.title || item.title;
            item.situation = data.situation || item.situation;
            item.text = item.situation;
            item.risk = data.risk || item.risk;
            item.proposal = data.proposal || item.proposal;
            saveState();
            renderExtraction();
            showToast("Redacción actualizada con éxito.", "success");
        } else {
            showToast("No se pudo redactar con IA.", "warning");
        }
    } catch (err) {
        console.error(err);
        showToast("Error al conectar con la IA.", "error");
    }
}

async function draftAllExtractionsWithAI() {
    if (!state.extracted.length) {
        showToast("No hay elementos para redactar.", "warning");
        return;
    }
    showToast("Redactando hallazgos candidatos con IA…", "info");
    for (let i = 0; i < state.extracted.length; i++) {
        await draftExtractionItemWithAI(i);
    }
    showToast("Todos los candidatos fueron redactados.", "success");
}

function findingFromItem(item) {
    return {
        id: crypto.randomUUID(),
        // Trazabilidad: relación directa con el candidato de origen
        sourceItemId: item.id,
        sourceItemIds: (item.sourceItemIds && item.sourceItemIds.length) ? item.sourceItemIds : [item.id],
        // Copia exacta de lo que el auditor aprobó — sin reinterpretación
        title: item.title || `${item.category}${item.originName ? ` - ${item.originName}` : ""}`,
        situation: item.situation || item.text || "",
        risk: item.risk || "",
        proposal: item.proposal || "",
        responsibleArea: "",
        actionOwner: "",
        severity: "Media",
        status: "Pendiente",
        targetDate: "",
        quantitativeBasis: "",
        sourceFile: item.filename || "",
        sourceLocation: item.originName || "",
        evidence: item.reference || "",
        ticket: "",
        followUp: ""
    };
}

// Convierte un candidato en hallazgo usando exactamente los datos que el auditor aprobó.
// Busca por ID (nunca por índice). No llama a IA automáticamente.
function convertOneToFinding(itemId) {
    const item = state.extracted.find(i => i.id === itemId);
    if (!item || item.converted || !isFindingEligible(item)) return;

    state.findings.push(findingFromItem(item));
    item.converted = true;
    item.selectedAsFinding = true;
    saveState();
    renderExtraction();
    renderFindings();
    renderMemoPreview();
    renderValidation();
    showToast("Hallazgo creado a partir del candidato seleccionado.", "success");
}


// Convierte en hallazgos solo los candidatos marcados con selectedAsFinding por el auditor.
function convertSelectedToFindings() {
    let created = 0;
    eligibleSelectedFindings(state.extracted).forEach(item => {
        state.findings.push(findingFromItem(item));
        item.converted = true;
        created += 1;
    });
    saveState();
    renderExtraction();
    renderFindings();
    renderMemoPreview();
    goToStep(3);
    showToast(created ? `${created} hallazgo(s) creados.` : "No había candidatos marcados como hallazgo.", "success");
}

// ============================================================
// HALLAZGOS
// ============================================================

function addFinding() {
    state.findings.push({
        id: crypto.randomUUID(), title: "", situation: "", risk: "", proposal: "",
        responsibleArea: "", actionOwner: "", severity: "Media", status: "Pendiente",
        targetDate: "", quantitativeBasis: "", sourceFile: "", sourceLocation: "",
        evidence: "", ticket: "", followUp: "", sourceItemId: null
    });
    saveState();
    renderFindings();
}

function updateFinding(index, field, value) {
    if (!state.findings[index]) return;
    state.findings[index][field] = value;
    scheduleSave();
    renderMemoPreview();
    renderValidation();
}

function deleteFinding(index) {
    const finding = state.findings[index];
    if (!finding) return;
    if (!window.confirm(`¿Eliminar el Hallazgo ${String(index + 1).padStart(2, "0")}?`)) return;
    if (finding.sourceItemId) {
        const source = state.extracted.find(item => item.id === finding.sourceItemId);
        if (source) {
            // Habilitar el candidato para ser convertido nuevamente
            source.converted = false;
            // selectedAsFinding se preserva en true: el auditor ya lo eligió como hallazgo
        }
    }
    state.findings.splice(index, 1);
    saveState();
    renderExtraction();
    renderFindings();
    renderMemoPreview();
    renderValidation();
}

function moveFinding(index, direction) {
    const target = index + direction;
    if (target < 0 || target >= state.findings.length) return;
    [state.findings[index], state.findings[target]] = [state.findings[target], state.findings[index]];
    saveState();
    renderFindings();
    renderMemoPreview();
}

function renderFindings() {
    const empty = el("findingsEmpty");
    const list = el("findingsList");
    if (!empty || !list) return;
    if (!state.findings.length) {
        empty.style.display = "block";
        list.innerHTML = "";
        return;
    }
    empty.style.display = "none";
    list.innerHTML = state.findings.map((finding, index) => `
        <article class="finding-card severity-${escapeHtml((finding.severity || "Media").toLowerCase())}">
            <div class="finding-header">
                <div><span class="finding-number">Hallazgo ${String(index + 1).padStart(2, "0")}</span><span class="severity-pill severity-${escapeHtml((finding.severity || "Media").toLowerCase())}">${escapeHtml(finding.severity || "Media")}</span></div>
                <div class="finding-actions">
                    <button type="button" class="icon-button" onclick="moveFinding(${index}, -1)">↑</button>
                    <button type="button" class="icon-button" onclick="moveFinding(${index}, 1)">↓</button>
                    <button type="button" class="icon-button danger" onclick="deleteFinding(${index})">×</button>
                </div>
            </div>
            ${finding.sourceFile ? `<div class="finding-source">Origen: <strong>${escapeHtml(finding.sourceFile)}</strong>${finding.sourceLocation ? ` · Solapa: ${escapeHtml(finding.sourceLocation)}` : ""}</div>` : ""}
            <div class="field field-wide">
                <div class="field-label-row">
                    <label>Título</label>
                    <button type="button" class="ai-button" onclick="improveFindingField(${index}, 'title', 'Título del hallazgo')">✦ Mejorar con IA</button>
                </div>
                <input value="${escapeHtml(finding.title)}" oninput="updateFinding(${index}, 'title', this.value)">
            </div>
            <div class="field field-wide">
                <div class="field-label-row">
                    <label>Situación observada</label>
                    <button type="button" class="ai-button" onclick="improveFindingField(${index}, 'situation', 'Situación observada')">✦ Mejorar con IA</button>
                </div>
                <textarea rows="4" oninput="updateFinding(${index}, 'situation', this.value)">${escapeHtml(finding.situation)}</textarea>
            </div>
            <div class="form-grid">
                <div class="field">
                    <div class="field-label-row">
                        <label>Riesgo</label>
                        <button type="button" class="ai-button" onclick="improveFindingField(${index}, 'risk', 'Riesgo')">✦ Mejorar con IA</button>
                    </div>
                    <textarea rows="4" oninput="updateFinding(${index}, 'risk', this.value)">${escapeHtml(finding.risk)}</textarea>
                </div>
                <div class="field">
                    <div class="field-label-row">
                        <label>Propuesta de mejora</label>
                        <button type="button" class="ai-button" onclick="improveFindingField(${index}, 'proposal', 'Propuesta de mejora')">✦ Mejorar con IA</button>
                    </div>
                    <textarea rows="4" oninput="updateFinding(${index}, 'proposal', this.value)">${escapeHtml(finding.proposal)}</textarea>
                </div>
                <div class="field"><label>Área responsable</label><input value="${escapeHtml(finding.responsibleArea)}" oninput="updateFinding(${index}, 'responsibleArea', this.value)"></div>
                <div class="field"><label>Criticidad</label><select onchange="updateFinding(${index}, 'severity', this.value); renderFindings();"><option ${finding.severity === "Alta" ? "selected" : ""}>Alta</option><option ${finding.severity === "Media" ? "selected" : ""}>Media</option><option ${finding.severity === "Baja" ? "selected" : ""}>Baja</option></select></div>
                <div class="field"><label>Estado</label><input value="${escapeHtml(finding.status)}" oninput="updateFinding(${index}, 'status', this.value)"></div>
                <div class="field"><label>Fecha compromiso</label><input type="date" value="${escapeHtml(finding.targetDate)}" oninput="updateFinding(${index}, 'targetDate', this.value)"></div>
            </div>
            <details class="finding-details"><summary>Seguimiento y trazabilidad</summary>
                <div class="form-grid">
                    <div class="field"><label>Responsable del plan</label><input value="${escapeHtml(finding.actionOwner)}" oninput="updateFinding(${index}, 'actionOwner', this.value)"></div>
                    <div class="field"><label>Base cuantitativa</label><input value="${escapeHtml(finding.quantitativeBasis)}" oninput="updateFinding(${index}, 'quantitativeBasis', this.value)"></div>
                    <div class="field"><label>Archivo de origen</label><input value="${escapeHtml(finding.sourceFile)}" oninput="updateFinding(${index}, 'sourceFile', this.value)"></div>
                    <div class="field"><label>Solapa / origen</label><input value="${escapeHtml(finding.sourceLocation)}" oninput="updateFinding(${index}, 'sourceLocation', this.value)"></div>
                    <div class="field"><label>Evidencia / referencia</label><input value="${escapeHtml(finding.evidence)}" oninput="updateFinding(${index}, 'evidence', this.value)"></div>
                    <div class="field"><label>Ticket</label><input value="${escapeHtml(finding.ticket)}" oninput="updateFinding(${index}, 'ticket', this.value)"></div>
                </div>
                <div class="field field-wide"><label>Seguimiento</label><textarea rows="3" oninput="updateFinding(${index}, 'followUp', this.value)">${escapeHtml(finding.followUp)}</textarea></div>
            </details>

        </article>
    `).join("");
}

// ============================================================
// MEMO PREVIEW
// ============================================================

function uniqueIncludedTexts(categories) {
    const seen = new Set();
    return state.extracted
        .filter(item => item.included && categories.has(item.category))
        .map(item => item.text || "")
        .filter(text => {
            const key = text.trim().toLowerCase();
            if (!key || seen.has(key)) return false;
            seen.add(key);
            return true;
        });
}

function severityBadgeHtml(severity) {
    const s = String(severity || "Media").toLowerCase();
    let bg = "#FFF4CC", color = "#8A6200";
    if (s === "alta") { bg = "#FDECEC"; color = "#B42318"; }
    else if (s === "baja") { bg = "#EAF6EA"; color = "#2E7D32"; }
    return `<span style="background:${bg};color:${color};padding:4px 8px;border-radius:12px;font-weight:700;font-size:11px;display:inline-block">${escapeHtml(severity || "Media")}</span>`;
}

function renderMemoPreview() {
    const container = el("memoPreview");
    if (!container) return;
    const tasks = uniqueIncludedTexts(new Set(["Tarea realizada"]));
    const results = uniqueIncludedTexts(new Set(["Conclusión", "Resultado", "Diferencia", "Observación", "Incumplimiento", "Pendiente", "Comentario"]));

    const findingRows = state.findings.length ? `
        <div style="overflow:auto;margin-top:12px"><table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr>${["N°","Título","Situación observada","Riesgo","Propuesta de mejora","Criticidad","Estado"].map(h => `<th style="background:#003DA5;color:white;padding:10px;border:1px solid #d0d7de;text-align:left">${h}</th>`).join("")}</tr></thead>
            <tbody>${state.findings.map((f, i) => `<tr>
                <td style="padding:10px;border:1px solid #d0d7de;font-weight:bold;text-align:center">${String(i+1).padStart(2,"0")}</td>
                <td style="padding:10px;border:1px solid #d0d7de;font-weight:bold">${escapeHtml(f.title)}</td>
                <td style="padding:10px;border:1px solid #d0d7de">${escapeHtml(f.situation)}</td>
                <td style="padding:10px;border:1px solid #d0d7de">${escapeHtml(f.risk)}</td>
                <td style="padding:10px;border:1px solid #d0d7de;background:#F9FAFB">${escapeHtml(f.proposal)}</td>
                <td style="padding:10px;border:1px solid #d0d7de;text-align:center">${severityBadgeHtml(f.severity)}</td>
                <td style="padding:10px;border:1px solid #d0d7de;text-align:center">${escapeHtml(f.status)}</td>
            </tr>`).join("")}</tbody>
        </table></div>` : `<p class="muted">No se incorporaron hallazgos.</p>`;

    const detailedFollowUp = state.findings.length ? `
        <div style="margin-top:24px">
            <h4 style="color:#003DA5;margin-bottom:12px">Seguimiento y Plan de Acción</h4>
            ${state.findings.map((f, i) => `
                <div style="background:#F3F7FB;border:1px solid #D0D7DE;border-radius:6px;padding:14px;margin-bottom:12px">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
                        <strong style="color:#003DA5">Hallazgo ${String(i+1).padStart(2,"0")}: ${escapeHtml(f.title)}</strong>
                        ${severityBadgeHtml(f.severity)}
                    </div>
                    <p style="margin:4px 0"><strong>Propuesta de mejora:</strong> ${escapeHtml(f.proposal || "Pendiente de definir")}</p>
                    <p style="margin:4px 0"><strong>Área responsable:</strong> ${escapeHtml(f.responsibleArea || "-")} &nbsp;|&nbsp; <strong>Responsable del plan:</strong> ${escapeHtml(f.actionOwner || "-")}</p>
                    <p style="margin:4px 0"><strong>Fecha compromiso:</strong> ${escapeHtml(f.targetDate || "-")} &nbsp;|&nbsp; <strong>Estado:</strong> ${escapeHtml(f.status || "Pendiente")}</p>
                    ${f.followUp ? `<p style="margin:4px 0;font-style:italic"><strong>Seguimiento:</strong> ${escapeHtml(f.followUp)}</p>` : ""}
                </div>
            `).join("")}
        </div>
    ` : "";

    container.innerHTML = `
        <div style="background:#003DA5;color:white;padding:22px 24px;border-radius:8px 8px 0 0"><h2 style="margin:0">MEMO – ${escapeHtml((state.general.title || "Auditoría").toUpperCase())}</h2></div>
        <div style="padding:22px">
            <p><strong>Área:</strong> ${escapeHtml(state.general.area)} &nbsp;&nbsp; <strong>Proceso:</strong> ${escapeHtml(state.general.process)}</p>
            <p><strong>Período:</strong> ${escapeHtml(state.general.period)} &nbsp;&nbsp; <strong>Auditor:</strong> ${escapeHtml(state.general.auditor)}</p>
            <h3 style="color:#003DA5">Objetivo</h3><p>${escapeHtml(state.general.objective)}</p>
            ${state.general.scope ? `<h3 style="color:#003DA5">Alcance</h3><p>${escapeHtml(state.general.scope)}</p>` : ""}
            <h3 style="color:#003DA5">Trabajo realizado</h3>${tasks.length ? `<ol>${tasks.map(t => `<li>${escapeHtml(t)}</li>`).join("")}</ol>` : `<p class="muted">Seleccioná tareas en Extracción para incorporarlas.</p>`}
            ${results.length ? `<h3 style="color:#003DA5">Resultados y observaciones relevantes</h3><ul>${results.map(r => `<li>${escapeHtml(r)}</li>`).join("")}</ul>` : ""}
            <h3 style="color:#003DA5">Hallazgos</h3>${findingRows}
            ${detailedFollowUp}
        </div>`;
}


// ============================================================
// VALIDACIONES / EXPORTACIÓN
// ============================================================

function foreignFindingSources() {
    const active = new Set((state.sources || []).map(s => (s.name || "").trim()).filter(Boolean));
    if (!active.size) return [];
    return state.findings.filter(f => (f.sourceFile || "").trim() && !active.has((f.sourceFile || "").trim()));
}

function validationIssues() {
    const issues = [];
    if (!state.general.title.trim()) issues.push("Falta el nombre de la auditoría.");
    if (!state.general.objective.trim()) issues.push("Falta completar el objetivo.");
    state.findings.forEach((f, i) => {
        if (!f.title.trim()) issues.push(`Hallazgo ${String(i+1).padStart(2,"0")}: falta título.`);
        if (!f.situation.trim() && !f.risk.trim()) issues.push(`Hallazgo ${String(i+1).padStart(2,"0")}: falta describir la situación observada.`);
    });
    const foreign = foreignFindingSources();
    if (foreign.length) issues.push(`${foreign.length} hallazgo(s) pertenecen a archivos que no forman parte de esta auditoría.`);
    return issues;
}

function renderValidation() {
    const container = el("validationList");
    if (!container) return;
    const issues = validationIssues();
    container.innerHTML = issues.length
        ? issues.map(issue => `<div class="validation-item warning">⚠ ${escapeHtml(issue)}</div>`).join("")
        : `<div class="validation-item success">✓ El memo no presenta validaciones bloqueantes.</div>`;
}

async function exportExcel() {
    Object.entries(GENERAL_FIELDS).forEach(([id, key]) => {
        if (el(id)) state.general[key] = el(id).value || "";
    });

    const foreign = foreignFindingSources();
    if (foreign.length) {
        showToast("Exportación bloqueada: hay hallazgos asociados a fuentes de otra auditoría.", "error");
        goToStep(5);
        return;
    }

    const issues = validationIssues();
    if (issues.length && !window.confirm(`El memo tiene validaciones pendientes:\n\n${issues.join("\n")}\n\n¿Querés exportarlo igualmente?`)) return;

    const button = document.querySelector('button[onclick*="exportExcel"]');
    if (button) { button.disabled = true; button.dataset.originalText = button.innerHTML; button.innerHTML = "Generando Excel…"; }

    try {
        const response = await fetch("/export-excel", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ memo: { auditId: state.auditId, general: state.general, findings: state.findings, sources: state.sources, extracted: state.extracted } })
        });
        if (!response.ok) {
            let data = {};
            try { data = await response.json(); } catch (_) {}
            throw new Error(data.error || "No se pudo generar el Excel.");
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `${(state.general.title || "Audit_Memo").replace(/[^a-z0-9áéíóúñü _-]/gi, "").trim().replace(/\s+/g, "_")}.xlsx`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        showToast("Excel generado correctamente.", "success");
    } catch (error) {
        console.error(error);
        showToast(error.message || "No se pudo exportar el memo.", "error");
    } finally {
        if (button) { button.disabled = false; if (button.dataset.originalText) button.innerHTML = button.dataset.originalText; }
    }
}

// ============================================================
// IA
// ============================================================

async function improveField(elementId, fieldType) {
    const input = el(elementId);
    if (!input || !input.value.trim()) {
        showToast("Primero escribí una idea para mejorar.", "warning");
        return;
    }
    const buttons = document.querySelectorAll(".ai-button");
    buttons.forEach(b => b.disabled = true);
    input.disabled = true;
    try {
        const response = await fetch("/improve-text", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: input.value.trim(), fieldType })
        });
        let data = {};
        try { data = await response.json(); } catch (_) {}
        if (!response.ok) throw new Error(data.error || "No se pudo mejorar el texto.");
        input.value = data.improved || input.value;
        input.dispatchEvent(new Event("input", { bubbles: true }));
        showToast("Redacción actualizada.", "success");
    } catch (error) {
        showToast(error.message || "No se pudo utilizar la IA.", "error");
    } finally {
        input.disabled = false;
        buttons.forEach(b => b.disabled = false);
    }
}

async function improveFindingField(findingIndex, fieldName, fieldTypeLabel) {

    const finding = state.findings[findingIndex];
    if (!finding) return;
    const currentVal = (finding[fieldName] || "").trim();
    if (!currentVal) {
        showToast(`Primero escribí una idea en ${fieldTypeLabel} para mejorar.`, "warning");
        return;
    }

    const buttons = document.querySelectorAll(".ai-button");
    buttons.forEach(b => b.disabled = true);
    try {
        const response = await fetch("/improve-text", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: currentVal, fieldType: fieldTypeLabel })
        });
        let data = {};
        try { data = await response.json(); } catch (_) {}
        if (!response.ok) throw new Error(data.error || "No se pudo mejorar el texto.");
        finding[fieldName] = data.improved || currentVal;
        saveState();
        renderFindings();
        renderMemoPreview();
        showToast("Redacción del hallazgo actualizada.", "success");
    } catch (error) {
        showToast(error.message || "No se pudo utilizar la IA.", "error");
    } finally {
        buttons.forEach(b => b.disabled = false);
    }
}


// ============================================================
// TOAST
// ============================================================

function showToast(message, type = "info") {
    const toast = el("toast");
    if (!toast) return;
    toast.textContent = message;
    toast.className = `toast show toast-${type}`;
    clearTimeout(showToast.timer);
    showToast.timer = setTimeout(() => { toast.className = "toast"; }, 4500);
}

// ============================================================
// INICIO
// ============================================================

if (typeof document !== "undefined") document.addEventListener("DOMContentLoaded", () => {
    loadState();
    hydrateGeneral();
    bindGeneralFields();
    setupDropZone();
    renderFiles();
    renderExtraction();
    renderFindings();
    renderMemoPreview();
    renderValidation();
    goToStep(state.currentStep || 1);
});

// ============================================================
// BASE DE CONOCIMIENTO Y MEMORIA DE AUDITORÍA
// ============================================================

let allKnowledgeItems = [];

async function toggleKnowledgeModal() {
    const modal = el("knowledgeModal");
    if (!modal) return;
    const isHidden = modal.style.display === "none" || !modal.style.display;
    modal.style.display = isHidden ? "flex" : "none";
    if (isHidden) {
        await loadKnowledgeItems();
    }
}

async function loadKnowledgeItems() {
    try {
        const response = await fetch("/knowledge");
        if (!response.ok) return;
        const data = await response.json();
        allKnowledgeItems = data.items || [];
        renderKnowledgeList(allKnowledgeItems);
    } catch (err) {
        console.error("Error cargando memoria de auditoría:", err);
    }
}

function renderKnowledgeList(items) {
    const list = el("knowledgeList");
    if (!list) return;
    if (!items.length) {
        list.innerHTML = `<div style="text-align: center; color: #64748b; padding: 20px;">No se encontraron elementos en la memoria.</div>`;
        return;
    }
    list.innerHTML = items.map(item => `
        <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px 16px;">
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 6px;">
                <div>
                    <span class="category-badge" style="background: #e2e8f0; color: #334155;">${escapeHtml(item.category)}</span>
                    <strong style="color: #0f172a; font-size: 14px; margin-left: 6px;">${escapeHtml(item.title)}</strong>
                </div>
                <button type="button" class="icon-button" onclick="deleteKnowledgeItem('${item.id}')" title="Eliminar regla">×</button>
            </div>
            <div style="font-size: 13px; color: #334155; margin-bottom: 4px;"><strong>Situación:</strong> ${escapeHtml(item.situation)}</div>
            ${item.risk ? `<div style="font-size: 12px; color: #64748b; margin-bottom: 2px;"><strong>Riesgo:</strong> ${escapeHtml(item.risk)}</div>` : ""}
            ${item.proposal ? `<div style="font-size: 12px; color: #64748b;"><strong>Propuesta:</strong> ${escapeHtml(item.proposal)}</div>` : ""}
            <div style="font-size: 11px; color: #94a3b8; margin-top: 6px;">Origen: ${escapeHtml(item.source_audit || 'Auditoría')}</div>
        </div>
    `).join("");
}

function filterKnowledgeItems(query) {
    const q = (query || "").toLowerCase().trim();
    if (!q) {
        renderKnowledgeList(allKnowledgeItems);
        return;
    }
    const filtered = allKnowledgeItems.filter(item =>
        (item.title || "").toLowerCase().includes(q) ||
        (item.situation || "").toLowerCase().includes(q) ||
        (item.category || "").toLowerCase().includes(q)
    );
    renderKnowledgeList(filtered);
}

function toggleAddRuleForm() {
    const form = el("addRuleForm");
    if (!form) return;
    form.style.display = form.style.display === "none" ? "block" : "none";
}

async function saveCustomRule() {
    const category = (el("newRuleCategory")?.value || "").trim();
    const title = (el("newRuleTitle")?.value || "").trim();
    const situation = (el("newRuleSituation")?.value || "").trim();
    const risk = (el("newRuleRisk")?.value || "").trim();
    const proposal = (el("newRuleProposal")?.value || "").trim();

    if (!title || !situation) {
        showToast("Completá al menos el Título y la Situación observada.", "warning");
        return;
    }

    try {
        const response = await fetch("/knowledge", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ category: category || "Regla de Firma", title, situation, risk, proposal })
        });
        if (response.ok) {
            showToast("Regla guardada en la memoria.", "success");
            if (el("newRuleCategory")) el("newRuleCategory").value = "";
            if (el("newRuleTitle")) el("newRuleTitle").value = "";
            if (el("newRuleSituation")) el("newRuleSituation").value = "";
            if (el("newRuleRisk")) el("newRuleRisk").value = "";
            if (el("newRuleProposal")) el("newRuleProposal").value = "";
            toggleAddRuleForm();
            await loadKnowledgeItems();
        } else {
            showToast("No se pudo guardar la regla.", "error");
        }
    } catch (err) {
        console.error(err);
        showToast("Error al guardar regla.", "error");
    }
}

async function deleteKnowledgeItem(itemId) {
    if (!confirm("¿Eliminar este elemento de la memoria de auditoría?")) return;
    try {
        const response = await fetch(`/knowledge/${itemId}`, { method: "DELETE" });
        if (response.ok) {
            showToast("Elemento eliminado de la memoria.", "success");
            await loadKnowledgeItems();
        }
    } catch (err) {
        console.error(err);
    }
}

if (typeof window !== "undefined") {
window.goToStep = goToStep;
window.extractInformation = extractInformation;
window.convertSelectedToFindings = convertSelectedToFindings;
window.addFinding = addFinding;
window.updateFinding = updateFinding;
window.deleteFinding = deleteFinding;
window.moveFinding = moveFinding;
window.toggleExtraction = toggleExtraction;
window.toggleSelectedAsFinding = toggleSelectedAsFinding;
window.updateExtractionText = updateExtractionText;
window.convertOneToFinding = convertOneToFinding;
window.removeFile = removeFile;
window.exportExcel = exportExcel;
window.improveField = improveField;
window.improveFindingField = improveFindingField;
window.startNewAudit = startNewAudit;
window.toggleKnowledgeModal = toggleKnowledgeModal;
window.filterKnowledgeItems = filterKnowledgeItems;
window.toggleAddRuleForm = toggleAddRuleForm;
window.saveCustomRule = saveCustomRule;
window.deleteKnowledgeItem = deleteKnowledgeItem;
window.draftAllExtractionsWithAI = draftAllExtractionsWithAI;

}

if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        onlyHallazgos,
        indexHallazgos,
        isFindingEligible,
        eligibleIncludedHallazgos,
        eligibleSelectedFindings,
        findingFromItem,
        getSourceItemForFinding,
        reconcileExtracted,
    };
}
