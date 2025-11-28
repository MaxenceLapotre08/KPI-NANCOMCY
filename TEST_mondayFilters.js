/**
 * TEST_mondayFilters.js
 * Tests progressifs pour valider les filtres Monday GraphQL query_params
 * 
 * APPROCHE:
 * 1. Test basique: 1 seul filtre Source
 * 2. Test: 2 filtres Source + Type
 * 3. Test: 3 filtres Source + Type + Statut
 * 4. Validation complète avant intégration
 */

// ==================== TEST 1: UN SEUL FILTRE ====================
function TEST_monday_1_filter() {
    Logger.log('========== TEST 1: UN SEUL FILTRE (Source) ==========');

    const boardId = SITE_MONDAY_LEADS_BOARD_ID;
    const sourceId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_SOURCE);

    Logger.log(`Board ID: ${boardId}`);
    Logger.log(`Source Column ID: ${sourceId}`);
    Logger.log(`Source Value: ${JSON.stringify(SITE_MONDAY_LEADS_SOURCE_MATCH)}`);

    // ✅ Test avec query_params - 1 seul filtre
    const queryParams = {
        rules: [
            {
                column_id: sourceId,
                compare_value: SITE_MONDAY_LEADS_SOURCE_MATCH,
                operator: "any_of"
            }
        ],
        operator: "and"
    };

    const q = `
    query($bid:[ID!], $qp:ItemsQuery){
      boards(ids:$bid){
        items_page(limit:50, query_params:$qp){
          cursor
          items{
            id name
            column_values(ids:["${sourceId}"]){ id text }
          }
        }
      }
    }`;

    try {
        Logger.log('\n📤 Envoi requête GraphQL...');
        Logger.log('Query params:', JSON.stringify(queryParams, null, 2));

        const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)], qp: queryParams });
        const page = d && d.boards && d.boards[0] && d.boards[0].items_page;

        if (!page) {
            Logger.log('❌ ERREUR: Pas de résultat');
            return { success: false, error: 'No page returned' };
        }

        const items = page.items || [];
        Logger.log(`\n✅ SUCCÈS: ${items.length} items récupérés avec filtre Source`);

        // Afficher les 5 premiers
        items.slice(0, 5).forEach((it, idx) => {
            const srcVal = (it.column_values.find(c => c.id === sourceId) || {}).text || '';
            Logger.log(`  [${idx + 1}] #${it.id} "${it.name}" - Source: "${srcVal}"`);
        });

        return { success: true, count: items.length };

    } catch (e) {
        Logger.log(`❌ ERREUR: ${e.message}`);
        return { success: false, error: e.message };
    }
}

// ==================== TEST 2: DEUX FILTRES ====================
function TEST_monday_2_filters() {
    Logger.log('\n========== TEST 2: DEUX FILTRES (Source + Type) ==========');

    const boardId = SITE_MONDAY_LEADS_BOARD_ID;
    const sourceId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_SOURCE);
    const typeId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_TYPE);

    Logger.log(`Source ID: ${sourceId} = ${JSON.stringify(SITE_MONDAY_LEADS_SOURCE_MATCH)}`);
    Logger.log(`Type ID: ${typeId} = ${JSON.stringify(SITE_MONDAY_LEADS_TYPE_MATCH)}`);

    // ✅ Test avec 2 filtres
    const queryParams = {
        rules: [
            {
                column_id: sourceId,
                compare_value: SITE_MONDAY_LEADS_SOURCE_MATCH,
                operator: "any_of"
            },
            {
                column_id: typeId,
                compare_value: SITE_MONDAY_LEADS_TYPE_MATCH,
                operator: "any_of"
            }
        ],
        operator: "and"
    };

    const q = `
    query($bid:[ID!], $qp:ItemsQuery){
      boards(ids:$bid){
        items_page(limit:50, query_params:$qp){
          cursor
          items{
            id name
            column_values(ids:["${sourceId}", "${typeId}"]){ id text }
          }
        }
      }
    }`;

    try {
        Logger.log('\n📤 Envoi requête GraphQL...');

        const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)], qp: queryParams });
        const page = d && d.boards && d.boards[0] && d.boards[0].items_page;

        if (!page) {
            Logger.log('❌ ERREUR: Pas de résultat');
            return { success: false, error: 'No page returned' };
        }

        const items = page.items || [];
        Logger.log(`\n✅ SUCCÈS: ${items.length} items récupérés avec filtres Source+Type`);

        // Afficher les 5 premiers
        items.slice(0, 5).forEach((it, idx) => {
            const srcVal = (it.column_values.find(c => c.id === sourceId) || {}).text || '';
            const typVal = (it.column_values.find(c => c.id === typeId) || {}).text || '';
            Logger.log(`  [${idx + 1}] #${it.id} "${it.name}" - Source: "${srcVal}" / Type: "${typVal}"`);
        });

        return { success: true, count: items.length };

    } catch (e) {
        Logger.log(`❌ ERREUR: ${e.message}`);
        return { success: false, error: e.message };
    }
}

// ==================== TEST 3: TROIS FILTRES ====================
function TEST_monday_3_filters() {
    Logger.log('\n========== TEST 3: TROIS FILTRES (Source + Type + Statut) ==========');

    const boardId = SITE_MONDAY_LEADS_BOARD_ID;
    const sourceId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_SOURCE);
    const typeId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_TYPE);
    const statusId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_STATUS);

    Logger.log(`Source ID: ${sourceId} = ${JSON.stringify(SITE_MONDAY_LEADS_SOURCE_MATCH)}`);
    Logger.log(`Type ID: ${typeId} = ${JSON.stringify(SITE_MONDAY_LEADS_TYPE_MATCH)}`);
    Logger.log(`Status ID: ${statusId} = ${JSON.stringify(SITE_MONDAY_LEADS_STATUS_MATCH)}`);

    // ✅ Test avec 3 filtres
    const queryParams = {
        rules: [
            {
                column_id: sourceId,
                compare_value: SITE_MONDAY_LEADS_SOURCE_MATCH,
                operator: "any_of"
            },
            {
                column_id: typeId,
                compare_value: SITE_MONDAY_LEADS_TYPE_MATCH,
                operator: "any_of"
            },
            {
                column_id: statusId,
                compare_value: SITE_MONDAY_LEADS_STATUS_MATCH,
                operator: "any_of"
            }
        ],
        operator: "and"
    };

    const q = `
    query($bid:[ID!], $qp:ItemsQuery){
      boards(ids:$bid){
        items_page(limit:50, query_params:$qp){
          cursor
          items{
            id name
            column_values(ids:["${sourceId}", "${typeId}", "${statusId}"]){ id text }
          }
        }
      }
    }`;

    try {
        Logger.log('\n📤 Envoi requête GraphQL...');

        const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)], qp: queryParams });
        const page = d && d.boards && d.boards[0] && d.boards[0].items_page;

        if (!page) {
            Logger.log('❌ ERREUR: Pas de résultat');
            return { success: false, error: 'No page returned' };
        }

        const items = page.items || [];
        Logger.log(`\n✅ SUCCÈS: ${items.length} items récupérés avec filtres Source+Type+Statut`);

        // Afficher tous
        items.forEach((it, idx) => {
            const srcVal = (it.column_values.find(c => c.id === sourceId) || {}).text || '';
            const typVal = (it.column_values.find(c => c.id === typeId) || {}).text || '';
            const stsVal = (it.column_values.find(c => c.id === statusId) || {}).text || '';
            Logger.log(`  [${idx + 1}] #${it.id} "${it.name}"`);
            Logger.log(`      Source: "${srcVal}" / Type: "${typVal}" / Statut: "${stsVal}"`);
        });

        return { success: true, count: items.length };

    } catch (e) {
        Logger.log(`❌ ERREUR: ${e.message}`);
        return { success: false, error: e.message };
    }
}

// ==================== TEST COMPLET ====================
function TEST_monday_ALL() {
    Logger.log('🧪 ========== LANCEMENT TESTS MONDAY FILTERS ========== 🧪\n');

    const results = {
        test1: TEST_monday_1_filter(),
        test2: TEST_monday_2_filters(),
        test3: TEST_monday_3_filters()
    };

    Logger.log('\n\n📊 ========== RÉSUMÉ DES TESTS ==========');
    Logger.log(`Test 1 (Source):            ${results.test1.success ? '✅ OK' : '❌ FAIL'} - ${results.test1.count || 0} items`);
    Logger.log(`Test 2 (Source+Type):       ${results.test2.success ? '✅ OK' : '❌ FAIL'} - ${results.test2.count || 0} items`);
    Logger.log(`Test 3 (Source+Type+Statut):${results.test3.success ? '✅ OK' : '❌ FAIL'} - ${results.test3.count || 0} items`);

    const allSuccess = results.test1.success && results.test2.success && results.test3.success;
    Logger.log(`\n${allSuccess ? '🎉 TOUS LES TESTS PASSÉS !' : '⚠️ CERTAINS TESTS ONT ÉCHOUÉ'}`);

    return { success: allSuccess, results };
}
