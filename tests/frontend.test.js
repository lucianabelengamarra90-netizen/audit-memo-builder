const test = require("node:test");
const assert = require("node:assert/strict");

if (!global.crypto) global.crypto = require("node:crypto").webcrypto;

const {
    onlyHallazgos,
    indexHallazgos,
    isFindingEligible,
    eligibleIncludedHallazgos,
} = require("../static/app.js");

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

