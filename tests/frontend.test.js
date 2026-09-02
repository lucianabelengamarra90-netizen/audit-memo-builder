const test = require("node:test");
const assert = require("node:assert/strict");

if (!global.crypto) global.crypto = require("node:crypto").webcrypto;

const {
    onlyHallazgos,
    indexHallazgos,
    isFindingEligible,
    eligibleIncludedHallazgos,
} = require("../static/app.js");

test("filters restored or received lists to Hallazgo only", () => {
    const hallazgo = { id: "h", category: "Hallazgo" };
    const items = [
        { id: "o", category: "Observación" },
        hallazgo,
        { id: "r", category: "Riesgo" },
    ];
    assert.deepEqual(onlyHallazgos(items), [hallazgo]);
    assert.deepEqual(onlyHallazgos(null), []);
});

test("keeps original indexes when preparing the rendered Hallazgo list", () => {
    const items = [
        { id: "old", category: "Diferencia" },
        { id: "first", category: "Hallazgo" },
        { id: "other", category: "Conclusión" },
        { id: "second", category: "Hallazgo" },
    ];
    assert.deepEqual(indexHallazgos(items).map(({ item, index }) => [item.id, index]), [
        ["first", 1], ["second", 3],
    ]);
});

test("only Hallazgo items are eligible conversion candidates", () => {
    const hallazgo = { category: "Hallazgo", included: true, converted: false };
    const converted = { category: "Hallazgo", included: true, converted: true };
    const risk = { category: "Riesgo", included: true, converted: false };
    assert.equal(isFindingEligible(hallazgo), true);
    assert.equal(isFindingEligible(risk), false);
    assert.deepEqual(eligibleIncludedHallazgos([risk, converted, hallazgo]), [hallazgo]);
});
