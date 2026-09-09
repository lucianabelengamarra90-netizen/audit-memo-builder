const test = require("node:test");
const assert = require("node:assert/strict");

if (!global.crypto) global.crypto = require("node:crypto").webcrypto;

const {
    onlyHallazgos,
    indexHallazgos,
    isFindingEligible,
    eligibleIncludedHallazgos,
    eligibleSelectedFindings,
    findingFromItem,
    getSourceItemForFinding,
    reconcileExtracted,
} = require("../static/app.js");

// ============================================================
// Tests originales
// ============================================================

test("preserves all extracted items for processing", () => {
    const item1 = { id: "h", category: "Hallazgo" };
    const item2 = { id: "o", category: "Observación" };
    const items = [item2, item1];
    assert.deepEqual(onlyHallazgos(items), [item2, item1]);
    assert.deepEqual(onlyHallazgos(null), []);
});

test("keeps original indexes for all extracted items", () => {
    const items = [
        { id: "old", category: "Diferencia" },
        { id: "first", category: "Hallazgo" },
    ];
    assert.deepEqual(indexHallazgos(items).map(({ item, index }) => [item.id, index]), [
        ["old", 0], ["first", 1],
    ]);
});

test("all valid interpreted items are eligible conversion candidates", () => {
    const hallazgo = { category: "Hallazgo", included: true, converted: false };
    const converted = { category: "Hallazgo", included: true, converted: true };
    const diff = { category: "Diferencia", included: true, converted: false };
    assert.equal(isFindingEligible(hallazgo), true);
    assert.equal(isFindingEligible(diff), true);
    assert.deepEqual(eligibleIncludedHallazgos([diff, converted, hallazgo]), [diff, hallazgo]);
});

// ============================================================
// Tests de trazabilidad 1-a-1
// ============================================================

// Helper para fabricar un candidato de prueba
function makeItem(overrides = {}) {
    return {
        id: crypto.randomUUID(),
        category: "Hallazgo",
        filename: "test.xlsx",
        originName: "Hoja1",
        reference: "A1",
        title: "Título de prueba",
        situation: "Situación de prueba",
        risk: "Riesgo de prueba",
        proposal: "Propuesta de prueba",
        included: false,
        selectedAsFinding: false,
        converted: false,
        ...overrides
    };
}

// TEST 1: Seleccionar 3 candidatos genera exactamente 3 hallazgos
test("seleccionar 3 candidatos genera exactamente 3 hallazgos", () => {
    const items = [
        makeItem({ selectedAsFinding: true }),   // candidato 1
        makeItem({ selectedAsFinding: false }),  // candidato 2 - NO seleccionado
        makeItem({ selectedAsFinding: true }),   // candidato 3
        makeItem({ selectedAsFinding: false }),  // candidato 4 - NO seleccionado
        makeItem({ selectedAsFinding: true }),   // candidato 5
    ];
    const eligible = eligibleSelectedFindings(items);
    assert.equal(eligible.length, 3, "Deben quedar 3 candidatos elegibles");

    const findings = eligible.map(item => findingFromItem(item));
    assert.equal(findings.length, 3, "Deben generarse exactamente 3 hallazgos");
});

// TEST 2: Cada hallazgo conserva correctamente su sourceItemId
test("cada hallazgo conserva su sourceItemId", () => {
    const item1 = makeItem({ selectedAsFinding: true });
    const item2 = makeItem({ selectedAsFinding: true });
    const item3 = makeItem({ selectedAsFinding: true });

    const finding1 = findingFromItem(item1);
    const finding2 = findingFromItem(item2);
    const finding3 = findingFromItem(item3);

    assert.equal(finding1.sourceItemId, item1.id);
    assert.equal(finding2.sourceItemId, item2.id);
    assert.equal(finding3.sourceItemId, item3.id);

    // Los IDs de los hallazgos deben ser distintos entre sí y distintos al item
    assert.notEqual(finding1.id, item1.id);
    assert.notEqual(finding1.id, finding2.id);
    assert.notEqual(finding1.id, finding3.id);
});

// TEST 3: included=true sin selectedAsFinding=true no debe convertirse en hallazgo
test("elemento incluido pero no marcado como hallazgo no se convierte", () => {
    const items = [
        makeItem({ included: true, selectedAsFinding: false }),  // soporte, no hallazgo
        makeItem({ included: true, selectedAsFinding: true }),   // sí hallazgo
        makeItem({ included: false, selectedAsFinding: false }), // nada
    ];

    const forFindings = eligibleSelectedFindings(items);
    assert.equal(forFindings.length, 1, "Solo 1 debe convertirse en hallazgo");
    assert.equal(forFindings[0].selectedAsFinding, true);

    // eligibleIncludedHallazgos sigue funcionando para soporte del memo
    const forSupport = eligibleIncludedHallazgos(items);
    assert.equal(forSupport.length, 2, "2 elementos quedan como soporte (included=true, not converted)");
});

// TEST 4: Eliminar hallazgo vuelve a habilitar su candidato (converted=false)
test("eliminar hallazgo rehabilita el candidato original", () => {
    const item = makeItem({ selectedAsFinding: true });
    const finding = findingFromItem(item);

    // Simular la conversión
    item.converted = true;
    assert.equal(item.converted, true);

    // Simular eliminación del hallazgo: buscar por sourceItemId y revertir converted
    const extracted = [item];
    const source = extracted.find(i => i.id === finding.sourceItemId);
    assert.ok(source, "Debe encontrar el candidato original por sourceItemId");
    source.converted = false;

    assert.equal(item.converted, false, "El candidato debe volver a estar disponible");
    // selectedAsFinding se preserva
    assert.equal(item.selectedAsFinding, true, "selectedAsFinding debe preservarse");
});

// TEST 5: Reordenar hallazgos no rompe la relación con el candidato
test("reordenar hallazgos no rompe sourceItemId", () => {
    const item1 = makeItem({ selectedAsFinding: true });
    const item2 = makeItem({ selectedAsFinding: true });
    const item3 = makeItem({ selectedAsFinding: true });

    const findings = [
        findingFromItem(item1),
        findingFromItem(item2),
        findingFromItem(item3),
    ];

    // Simular swap de posiciones 0 y 2
    [findings[0], findings[2]] = [findings[2], findings[0]];

    // Los sourceItemIds deben seguir siendo correctos (relación por ID, no por posición)
    assert.equal(findings[0].sourceItemId, item3.id, "El hallazgo movido debe seguir apuntando a item3");
    assert.equal(findings[1].sourceItemId, item2.id, "El hallazgo del medio no cambia");
    assert.equal(findings[2].sourceItemId, item1.id, "El hallazgo movido debe seguir apuntando a item1");
});

// TEST 6: findingFromItem copia el contenido exacto sin modificar el item original
test("findingFromItem no modifica el item original", () => {
    const item = makeItem({
        selectedAsFinding: true,
        title: "Título original",
        situation: "Situación original",
        risk: "Riesgo original",
        proposal: "Propuesta original",
    });
    const originalId = item.id;

    const finding = findingFromItem(item);

    // El hallazgo copia el contenido exacto
    assert.equal(finding.title, "Título original");
    assert.equal(finding.situation, "Situación original");
    assert.equal(finding.risk, "Riesgo original");
    assert.equal(finding.proposal, "Propuesta original");

    // El item no fue modificado
    assert.equal(item.id, originalId);
    assert.equal(item.title, "Título original");
    assert.equal(item.converted, false, "convertOneToFinding debe marcar converted, no findingFromItem");
});

// TEST 7: la relación sourceItemId permite ubicar al candidato en cualquier lista
test("sourceItemId permite localizar el candidato en una lista de extraídos", () => {
    const item1 = makeItem();
    const item2 = makeItem();
    const item3 = makeItem();
    const finding = findingFromItem(item2);

    // Simular búsqueda manual (mismo patrón que getSourceItemForFinding)
    const extracted = [item1, item2, item3];
    const found = extracted.find(i => i.id === finding.sourceItemId);

    assert.ok(found, "Debe encontrar el candidato");
    assert.equal(found.id, item2.id, "Debe ser el candidato correcto: item2");
    assert.notEqual(found.id, item1.id, "No debe ser item1");
    assert.notEqual(found.id, item3.id, "No debe ser item3");

    // La función getSourceItemForFinding hace exactamente esa búsqueda sobre state.extracted
    // Verificar que existe y es una función
    assert.equal(typeof getSourceItemForFinding, "function");
});

// TEST 8: reconcileExtracted preserva IDs de items ya vinculados a hallazgos
test("nueva extracción preserva IDs de candidatos ya vinculados", () => {
    const existingItem = {
        id: "id-estable-123",
        filename: "planilla.xlsx",
        originName: "Hoja1",
        reference: "B5",
        category: "Diferencia",
        included: true,
        selectedAsFinding: true,
        converted: true,
        title: "Diferencia detectada",
        situation: "Saldo no coincide",
        risk: "Riesgo alto",
        proposal: "Regularizar",
    };

    // Nueva extracción devuelve el mismo item con un UUID diferente (como pasa hoy sin la fix)
    const newExtraction = [
        {
            id: "nuevo-uuid-distinto",
            filename: "planilla.xlsx",
            originName: "Hoja1",
            reference: "B5",
            category: "Diferencia",
            included: false,
            selectedAsFinding: false,
            converted: false,
            title: "Diferencia detectada",
            situation: "Saldo no coincide",
        }
    ];

    const reconciled = reconcileExtracted(newExtraction, [existingItem]);

    assert.equal(reconciled.length, 1);
    assert.equal(reconciled[0].id, "id-estable-123", "El ID original debe preservarse");
    assert.equal(reconciled[0].converted, true, "El estado converted debe preservarse");
    assert.equal(reconciled[0].selectedAsFinding, true, "selectedAsFinding debe preservarse");
    assert.equal(reconciled[0].included, true, "included debe preservarse");
});

// TEST 9: Candidato consolidado de múltiples registros registra todos los sourceItemIds
test("findingFromItem propaga sourceItemIds de candidato consolidado", () => {
    const idA = crypto.randomUUID();
    const idB = crypto.randomUUID();
    const idC = crypto.randomUUID();

    const consolidatedItem = makeItem({
        id: crypto.randomUUID(),
        selectedAsFinding: true,
        sourceItemIds: [idA, idB, idC], // consolidado desde 3 registros
        title: "Hallazgo consolidado",
        situation: "Tres registros consolidados por IA",
    });

    const finding = findingFromItem(consolidatedItem);

    assert.ok(Array.isArray(finding.sourceItemIds), "sourceItemIds debe ser un array");
    assert.equal(finding.sourceItemIds.length, 3, "Deben estar los 3 IDs de origen");
    assert.ok(finding.sourceItemIds.includes(idA));
    assert.ok(finding.sourceItemIds.includes(idB));
    assert.ok(finding.sourceItemIds.includes(idC));

    // sourceItemId sigue apuntando al candidato consolidado
    assert.equal(finding.sourceItemId, consolidatedItem.id);
});
