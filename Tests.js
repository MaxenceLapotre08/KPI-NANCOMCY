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
 * Test de SheetHelpers.monthRangeKeys
 */
function test_SheetHelpers_monthRangeKeys() {
    const tests = [
        {
            name: "Même mois",
            from: new Date(2024, 0, 1),
            to: new Date(2024, 0, 31),
            expected: ["2024-01"]
        },
        {
            name: "3 mois consécutifs",
            from: new Date(2024, 0, 1),
            to: new Date(2024, 2, 31),
            expected: ["2024-01", "2024-02", "2024-03"]
        },
        {
            name: "À cheval sur 2 années",
            from: new Date(2024, 11, 1),
            to: new Date(2025, 1, 28),
            expected: ["2024-12", "2025-01", "2025-02"]
        }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        const result = SheetHelpers.monthRangeKeys(test.from, test.to);
        const resultStr = JSON.stringify(result);
        const expectedStr = JSON.stringify(test.expected);

        if (resultStr === expectedStr) {
            passed++;
        } else {
            failed++;
            Logger.log(`❌ FAIL [${test.name}]: SheetHelpers.monthRangeKeys = ${resultStr} (expected: ${expectedStr})`);
        }
    });

    Logger.log(`SheetHelpers.monthRangeKeys: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test de ErrorHandler.logError
 */
function test_ErrorHandler_logError() {
    const tests = [
        {
            name: "Erreur simple (string)",
            context: "TestContext",
            error: "Test error message",
            severity: "ERROR"
        },
        {
            name: "Erreur objet (Error)",
            context: "TestContext",
            error: new Error("Test error object"),
            severity: "WARNING"
        },
        {
            name: "Info avec metadata",
            context: "TestContext",
            error: "Info message",
            severity: "INFO"
        }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        try {
            const result = ErrorHandler.logError(test.context, test.error, test.severity, { test: true });

            // Vérifier que le logEntry a la bonne structure
            if (result && result.context === test.context && result.severity === test.severity) {
                passed++;
            } else {
                failed++;
                Logger.log(`❌ FAIL [${test.name}]: Structure logEntry invalide`);
            }
        } catch (e) {
            failed++;
            Logger.log(`❌ FAIL [${test.name}]: Exception levée: ${e.message}`);
        }
    });

    Logger.log(`ErrorHandler.logError: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test de ErrorHandler.wrapFunction
 */
function test_ErrorHandler_wrapFunction() {
    const tests = [
        {
            name: "Fonction qui réussit",
            fn: () => 42,
            expectedResult: 42,
            shouldThrow: false
        },
        {
            name: "Fonction qui échoue avec rethrow=true",
            fn: () => { throw new Error("Test error"); },
            expectedResult: null,
            shouldThrow: true
        },
        {
            name: "Fonction qui échoue avec rethrow=false",
            fn: () => { throw new Error("Test error"); },
            expectedResult: null,
            shouldThrow: false,
            options: { rethrow: false }
        }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        try {
            const result = ErrorHandler.wrapFunction(test.fn, "TestContext", test.options || {});

            if (test.shouldThrow && !test.options) {
                failed++;
                Logger.log(`❌ FAIL [${test.name}]: Devrait throw mais n'a pas throw`);
            } else if (result === test.expectedResult) {
                passed++;
            } else {
                failed++;
                Logger.log(`❌ FAIL [${test.name}]: Résultat ${result} != ${test.expectedResult}`);
            }
        } catch (e) {
            if (test.shouldThrow && (!test.options || test.options.rethrow !== false)) {
                passed++;
            } else {
                failed++;
                Logger.log(`❌ FAIL [${test.name}]: Ne devrait pas throw mais a throw: ${e.message}`);
            }
        }
    });

    Logger.log(`ErrorHandler.wrapFunction: ${passed} passed, ${failed} failed`);
    return failed === 0;
}

/**
 * Test des getters Config.js
 */
function test_Config_getters() {
    const tests = [
        {
            name: "getConfig avec clé valide",
            fn: () => getConfig('DEFAULT_RETRY_COUNT'),
            expected: 3
        },
        {
            name: "getConfig avec clé invalide (doit throw)",
            fn: () => getConfig('INVALID_KEY'),
            shouldThrow: true
        }
    ];

    let passed = 0;
    let failed = 0;

    tests.forEach(test => {
        try {
            const result = test.fn();

            if (test.shouldThrow) {
                failed++;
                Logger.log(`❌ FAIL [${test.name}]: Devrait throw mais n'a pas throw`);
            } else if (result === test.expected) {
                passed++;
            } else {
                failed++;
                Logger.log(`❌ FAIL [${test.name}]: Résultat ${result} != ${test.expected}`);
            }
        } catch (e) {
            if (test.shouldThrow) {
                passed++;
            } else {
                failed++;
                Logger.log(`❌ FAIL [${test.name}]: Exception inattendue: ${e.message}`);
            }
        }
    });

    Logger.log(`Config getters: ${passed} passed, ${failed} failed`);
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
        'SITE_parseMatomoVisitsResponse_': test_SITE_parseMatomoVisitsResponse_(),
        'SheetHelpers.monthRangeKeys': test_SheetHelpers_monthRangeKeys(),
        'ErrorHandler.logError': test_ErrorHandler_logError(),
        'ErrorHandler.wrapFunction': test_ErrorHandler_wrapFunction(),
        'Config getters': test_Config_getters()
    };

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);

    const passed = Object.values(results).filter(r => r === true).length;
    const failed = Object.values(results).filter(r => r === false).length;
    const total = Object.keys(results).length;

    Logger.log('========================================');
    Logger.log(`RÉSULTATS: ${passed}/${total} suites passed, ${failed} failed`);
    Logger.log(`Temps d'exécution: ${elapsed}s`);
    Logger.log(`Couverture estimée: ~${Math.round(total / 200 * 100)}% (${total} fonctions testées sur ~200)`);
    Logger.log('========================================');

    if (failed === 0) {
        Utils_toast(`✅ Tous les tests sont passés (${elapsed}s) - ${total} suites`, 'Tests', 5);
        return true;
    } else {
        Utils_toast(`❌ ${failed}/${total} tests échoués (voir logs)`, 'Tests', 10);
        throw new Error(`${failed} test suite(s) failed. Check logs for details.`);
    }
}
