/**
 * Tests.js
 * Tests unitaires pour valider le bon fonctionnement des fonctions critiques.
 * Exécuter: runAllTests()
 */

/**
 * Test de la fonction Utils_normHeader
 */
function test_Utils_normHeader() {
    const tests = [
        { input: "Budget Investi", expected: "budget investi" },
        { input: "Coût par Lead  ", expected: "cout par lead" },
        { input: "CTR", expected: "ctr" },
        { input: "Nombre d'Impressions", expected: "nombre dimpressions" },
        { input: "  Taux de Rebond  ", expected: "taux de rebond" },
        { input: "ROAS", expected: "roas" }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        const result = Utils_normHeader(test.input);
        if (result === test.expected) {
            passed++;
        } else {
            failed++;
            Logger.log(`❌ FAIL: Utils_normHeader("${test.input}") = "${result}" (expected: "${test.expected}")`);
        }
    });

    Logger.log(`Utils_normHeader: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test de la fonction Utils_sheetCellToYYYYMM
 */
function test_Utils_sheetCellToYYYYMM() {
    const tests = [
        { input: "novembre 2025", expected: "2025-11" },
        { input: "janv 2024", expected: "2024-01" },
        { input: new Date(2024, 0, 15), expected: "2024-01" },
        { input: "2024-03", expected: "2024-03" },
        { input: "2024/12", expected: "2024-12" },
        { input: "202405", expected: "2024-05" }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        const result = Utils_sheetCellToYYYYMM(test.input);
        if (result === test.expected) {
            passed++;
        } else {
            failed++;
            const inputStr = test.input instanceof Date ? test.input.toISOString() : test.input;
            Logger.log(`❌ FAIL: Utils_sheetCellToYYYYMM("${inputStr}") = "${result}" (expected: "${test.expected}")`);
        }
    });

    Logger.log(`Utils_sheetCellToYYYYMM: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test de la fonction Utils_monthKeyToFr
 */
function test_Utils_monthKeyToFr() {
    const tests = [
        { input: "2025-11", expected: "novembre 2025" },
        { input: "2024-01", expected: "janvier 2024" },
        { input: "2024-12", expected: "décembre 2024" }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        const result = Utils_monthKeyToFr(test.input);
        if (result === test.expected) {
            passed++;
        } else {
            failed++;
            Logger.log(`❌ FAIL: Utils_monthKeyToFr("${test.input}") = "${result}" (expected: "${test.expected}")`);
        }
    });

    Logger.log(`Utils_monthKeyToFr: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test de la fonction Utils_textEqualsAny
 */
function test_Utils_textEqualsAny() {
    const tests = [
        { input: ["Site Internet", ["site internet", "site-internet"]], expected: true },
        { input: ["Meta Ads", ["meta", "facebook"]], expected: false },
        { input: ["Lead", ["lead", "prospect"]], expected: true }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        const [txt, arr] = test.input;
        const result = Utils_textEqualsAny(txt, arr);
        if (result === test.expected) {
            passed++;
        } else {
            failed++;
            Logger.log(`❌ FAIL: Utils_textEqualsAny("${txt}", ${JSON.stringify(arr)}) = ${result} (expected: ${test.expected})`);
        }
    });

    Logger.log(`Utils_textEqualsAny: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test de la fonction Utils_textIncludesAny
 */
function test_Utils_textIncludesAny() {
    const tests = [
        { input: ["Site Internet", ["site", "internet"]], expected: true },
        { input: ["Formulaire", ["form", "formulaire"]], expected: true },
        { input: ["Appel", ["email", "message"]], expected: false }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        const [txt, arr] = test.input;
        const result = Utils_textIncludesAny(txt, arr);
        if (result === test.expected) {
            passed++;
        } else {
            failed++;
            Logger.log(`❌ FAIL: Utils_textIncludesAny("${txt}", ${JSON.stringify(arr)}) = ${result} (expected: ${test.expected})`);
        }
    });

    Logger.log(`Utils_textIncludesAny: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test de la fonction SITE_parseMatomoVisitsResponse_
 */
function test_SITE_parseMatomoVisitsResponse_() {
    const tests = [
        {
            name: "Réponse valide",
            input: { "2025-11": { nb_visits: 456, nb_actions: 1234 } },
            expected: { "2025-11": { sessions: 456 } }
        },
        {
            name: "Erreur Matomo (string)",
            input: { "2025-11": "No data available" },
            expected: {}
        },
        {
            name: "Objet vide",
            input: {},
            expected: {}
        },
        {
            name: "Null",
            input: null,
            expected: {}
        }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        const result = SITE_parseMatomoVisitsResponse_(test.input);
        const resultStr = JSON.stringify(result);
        const expectedStr = JSON.stringify(test.expected);

        if (resultStr === expectedStr) {
            passed++;
        } else {
            failed++;
            Logger.log(`❌ FAIL [${test.name}]: SITE_parseMatomoVisitsResponse_ = ${resultStr} (expected: ${expectedStr})`);
        }
    });

    Logger.log(`SITE_parseMatomoVisitsResponse_: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Exécute tous les tests
 */
function runAllTests() {
    Logger.log('========================================');
    Logger.log('   EXÉCUTION DES TESTS UNITAIRES');
    Logger.log('========================================');

    const startTime = Date.now();

    Utils_toast('Exécution des tests...', 'Tests', 3);

    const results = {
        'Utils_normHeader': test_Utils_normHeader(),
        'Utils_sheetCellToYYYYMM': test_Utils_sheetCellToYYYYMM(),
        'Utils_monthKeyToFr': test_Utils_monthKeyToFr(),
        'Utils_textEqualsAny': test_Utils_textEqualsAny(),
        'Utils_textIncludesAny': test_Utils_textIncludesAny(),
        'SITE_parseMatomoVisitsResponse_': test_SITE_parseMatomoVisitsResponse_()
    };

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);

    const passed = Object.values(results).filter(r => r === true).length;
    const failed = Object.values(results).filter(r => r === false).length;

    Logger.log('========================================');
    Logger.log(`RÉSULTATS: ${passed} suites passed, ${failed} suites failed`);
    Logger.log(`Temps d'exécution: ${elapsed}s`);
    Logger.log('========================================');

    if (failed === 0) {
        Utils_toast(`✅ Tous les tests sont passés (${elapsed}s)`, 'Tests', 5);
        return true;
    } else {
        Utils_toast(`❌ ${failed} suites de tests échouées (voir les logs)`, 'Tests', 10);
        throw new Error(`${failed} test suite(s) failed. Check logs for details.`);
    }
}
