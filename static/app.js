// ============================================================
// AUDIT MEMO BUILDER - FRONTEND ESTABLE
// ============================================================

const STORAGE_PREFIX = "auditMemoBuilder:v3";
const AUTOSAVE_DELAY = 350;
const FINDING_ELIGIBLE = new Set([
    "Hallazgo", "Observación", "Diferencia", "Incumplimiento",
    "Riesgo", "Pendiente", "Conclusión"
]);

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
            extracted: saved.extracted || [],
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
    renderExtractionStatus("Analizando todas las solapas del papel de trabajo…", "loading");
    if (el("extractionList")) el("extractionList").innerHTML = "";
    if (el("extractedCount")) el("extractedCount").textContent = "Procesando…";

    const form = new FormData();
    selectedFiles.forEach(file => form.append("files", file));
    form.append("freeText", freeText);

    try {
        const response = await fetch("/extract", { method: "POST", body: form });
        let data = {};
        try { data = await response.json(); } catch (_) {}
        if (!response.ok) throw new Error(data.error || "No se pudo procesar la documentación.");

        state.extracted = (data.items || []).map(item => ({
            ...item,
            included: Boolean(item.included),
            converted: Boolean(item.converted)
        }));
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

        const warningText = (data.warnings || []).length ? ` · ${(data.warnings || []).length} advertencia(s)` : "";
        showToast(`${data.message || "Extracción finalizada"}${warningText}`, (data.errors || []).length ? "warning" : "success");
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

function renderExtraction() {
    const empty = el("extractionEmpty");
    const list = el("extractionList");
    const count = el("extractedCount");
    if (!empty || !list) return;

    if (extractionInProgress) {
        renderExtractionStatus("Analizando todas las solapas del papel de trabajo…", "loading");
        return;
    }
    if (count) count.textContent = `${state.extracted.length} elementos`;
    if (!state.extracted.length) {
        empty.style.display = "block";
        empty.innerHTML = `<div class="empty-icon">⌕</div><h3>Todavía no hay información extraída</h3><p>Cargá tus papeles de trabajo desde el paso anterior.</p>`;
        list.innerHTML = "";
        return;
    }

    empty.style.display = "none";
    list.innerHTML = state.extracted.map((item, index) => {
        const eligible = FINDING_ELIGIBLE.has(item.category);
        return `
            <article class="extraction-card">
                <div class="extraction-card-header">
                    <div><span class="category-badge">${escapeHtml(item.category)}</span><strong class="source-title">${escapeHtml(item.filename)}</strong></div>
                    <label class="include-check"><input type="checkbox" ${item.included ? "checked" : ""} onchange="toggleExtraction(${index}, this.checked)"> Incluir</label>
                </div>
                <textarea class="extraction-text" oninput="updateExtractionText(${index}, this.value)">${escapeHtml(item.text)}</textarea>
                <div class="trace-meta">
                    ${item.originName ? `<span>Solapa: <strong>${escapeHtml(item.originName)}</strong></span>` : ""}
                    ${item.reference ? `<span>${escapeHtml(item.reference)}</span>` : ""}
                    ${item.keyword ? `<span>Detectado por: ${escapeHtml(item.keyword)}</span>` : ""}
                </div>
                <div class="extraction-card-footer">
                    <span>${item.converted ? "Convertido en hallazgo" : eligible ? "Puede convertirse en hallazgo" : "Se incorporará como soporte del memo"}</span>
                    ${eligible && !item.converted ? `<button type="button" class="btn btn-secondary" onclick="convertOneToFinding(${index})">Crear hallazgo</button>` : ""}
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
function updateExtractionText(index, value) {
    if (!state.extracted[index]) return;
    state.extracted[index].text = value;
    scheduleSave();
    renderMemoPreview();
}

function findingFromItem(item) {
    return {
        id: crypto.randomUUID(),
        title: `${item.category}${item.originName ? ` - ${item.originName}` : ""}`,
        situation: item.category === "Riesgo" ? "" : (item.text || ""),
        risk: item.category === "Riesgo" ? (item.text || "") : "",
        proposal: "",
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
        followUp: "",
        sourceItemId: item.id
    };
}

function convertOneToFinding(index) {
    const item = state.extracted[index];
    if (!item || item.converted || !FINDING_ELIGIBLE.has(item.category)) return;
    state.findings.push(findingFromItem(item));
    item.converted = true;
    item.included = true;
    saveState();
    renderExtraction();
    renderFindings();
    renderMemoPreview();
    renderValidation();
    showToast("Hallazgo creado para revisión.", "success");
}

function convertSelectedToFindings() {
    let created = 0;
    state.extracted.forEach(item => {
        if (item.included && !item.converted && FINDING_ELIGIBLE.has(item.category)) {
            state.findings.push(findingFromItem(item));
            item.converted = true;
            created += 1;
        }
    });
    saveState();
    renderExtraction();
    renderFindings();
    renderMemoPreview();
    goToStep(3);
    showToast(created ? `${created} hallazgo(s) creados para revisión.` : "Los elementos seleccionados quedaron incorporados al memo.", "success");
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
        if (source) source.converted = false;
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
            <div class="field field-wide"><label>Título</label><input value="${escapeHtml(finding.title)}" oninput="updateFinding(${index}, 'title', this.value)"></div>
            <div class="field field-wide"><label>Situación observada</label><textarea rows="4" oninput="updateFinding(${index}, 'situation', this.value)">${escapeHtml(finding.situation)}</textarea></div>
            <div class="form-grid">
                <div class="field"><label>Riesgo</label><textarea rows="4" oninput="updateFinding(${index}, 'risk', this.value)">${escapeHtml(finding.risk)}</textarea></div>
                <div class="field"><label>Propuesta de mejora</label><textarea rows="4" oninput="updateFinding(${index}, 'proposal', this.value)">${escapeHtml(finding.proposal)}</textarea></div>
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

function renderMemoPreview() {
    const container = el("memoPreview");
    if (!container) return;
    const tasks = uniqueIncludedTexts(new Set(["Tarea realizada"]));
    const results = uniqueIncludedTexts(new Set(["Conclusión", "Resultado", "Diferencia", "Observación", "Incumplimiento", "Pendiente", "Comentario"]));

    const findingRows = state.findings.length ? `
        <div style="overflow:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr>${["N°","Título","Situación observada","Riesgo","Propuesta de mejora","Criticidad","Estado"].map(h => `<th style="background:#17365D;color:white;padding:9px;border:1px solid #d0d7de;text-align:left">${h}</th>`).join("")}</tr></thead>
            <tbody>${state.findings.map((f, i) => `<tr>
                <td style="padding:8px;border:1px solid #d0d7de">${String(i+1).padStart(2,"0")}</td>
                <td style="padding:8px;border:1px solid #d0d7de">${escapeHtml(f.title)}</td>
                <td style="padding:8px;border:1px solid #d0d7de">${escapeHtml(f.situation)}</td>
                <td style="padding:8px;border:1px solid #d0d7de">${escapeHtml(f.risk)}</td>
                <td style="padding:8px;border:1px solid #d0d7de">${escapeHtml(f.proposal)}</td>
                <td style="padding:8px;border:1px solid #d0d7de;font-weight:700">${escapeHtml(f.severity)}</td>
                <td style="padding:8px;border:1px solid #d0d7de">${escapeHtml(f.status)}</td>
            </tr>`).join("")}</tbody>
        </table></div>` : `<p class="muted">No se incorporaron hallazgos.</p>`;

    container.innerHTML = `
        <div style="background:#17365D;color:white;padding:22px 24px;border-radius:8px 8px 0 0"><h2 style="margin:0">MEMO – ${escapeHtml((state.general.title || "Auditoría").toUpperCase())}</h2></div>
        <div style="padding:22px">
            <p><strong>Área:</strong> ${escapeHtml(state.general.area)} &nbsp;&nbsp; <strong>Proceso:</strong> ${escapeHtml(state.general.process)}</p>
            <p><strong>Período:</strong> ${escapeHtml(state.general.period)} &nbsp;&nbsp; <strong>Auditor:</strong> ${escapeHtml(state.general.auditor)}</p>
            <h3>Objetivo</h3><p>${escapeHtml(state.general.objective)}</p>
            ${state.general.scope ? `<h3>Alcance</h3><p>${escapeHtml(state.general.scope)}</p>` : ""}
            <h3>Trabajo realizado</h3>${tasks.length ? `<ol>${tasks.map(t => `<li>${escapeHtml(t)}</li>`).join("")}</ol>` : `<p class="muted">Seleccioná tareas en Extracción para incorporarlas.</p>`}
            ${results.length ? `<h3>Resultados y observaciones relevantes</h3><ul>${results.map(r => `<li>${escapeHtml(r)}</li>`).join("")}</ul>` : ""}
            <h3>Hallazgos</h3>${findingRows}
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

document.addEventListener("DOMContentLoaded", () => {
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

window.goToStep = goToStep;
window.extractInformation = extractInformation;
window.convertSelectedToFindings = convertSelectedToFindings;
window.addFinding = addFinding;
window.updateFinding = updateFinding;
window.deleteFinding = deleteFinding;
window.moveFinding = moveFinding;
window.toggleExtraction = toggleExtraction;
window.updateExtractionText = updateExtractionText;
window.convertOneToFinding = convertOneToFinding;
window.removeFile = removeFile;
window.exportExcel = exportExcel;
window.improveField = improveField;
window.startNewAudit = startNewAudit;
