/**
 * TEST_mondayFilters_INDEX.js
 * Tests avec les INDEX plutôt que les textes
 */

function TEST_monday_with_INDEX() {
    Logger.log('========== TEST FILTRES MONDAY AVEC INDEX ==========');

    const boardId = SITE_MONDAY_LEADS_BOARD_ID;
    const sourceId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_SOURCE);
    const typeId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_TYPE);
    const statusId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_STATUS);

    Logger.log('\n📋 Mapping trouvé:');
    Logger.log('  Source "SITE INTERNET" = index: 2');
    Logger.log('  Type "Appel" = index: 2');
    Logger.log('  Type "Formulaire" = index: 1');
    Logger.log('  Status "Lead" = index: 1');

    // ✅ FILTRES AVEC INDEX (au lieu de texte)
    const queryParams = {
        rules: [
            {
                column_id: sourceId,
                compare_value: [2],  // ✅ INDEX 2 = "SITE INTERNET"
                operator: "any_of"
            },
            {
                column_id: typeId,
                compare_value: [2],  // ✅ INDEX 2 = "Appel"
                operator: "any_of"
            },
            {
                column_id: statusId,
                compare_value: [1],  // ✅ INDEX 1 = "Lead"
                operator: "any_of"
            }
        ],
        operator: "and"
    };

    const q = `
    query($bid:[ID!], $qp:ItemsQuery){
      boards(ids:$bid){
        items_page(limit:100, query_params:$qp){
          cursor
          items{
            id name created_at
            column_values(ids:["${sourceId}", "${typeId}", "${statusId}"]){ 
              id text value
            }
          }
        }
      }
    }`;

    try {
        Logger.log('\n📤 Envoi requête avec INDEX...');
        Logger.log('Query params:', JSON.stringify(queryParams, null, 2));

        const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)], qp: queryParams });
        const page = d && d.boards && d.boards[0] && d.boards[0].items_page;

        if (!page) {
            Logger.log('❌ ERREUR: Pas de résultat');
            return { success: false, error: 'No page returned' };
        }

        const items = page.items || [];
        Logger.log(`\n✅ SUCCÈS: ${items.length} items récupérés !`);
        Logger.log(`🎯 Critères: Source="SITE INTERNET" + Type="Appel" + Status="Lead"\n`);

        // Afficher TOUS les items trouvés
        items.forEach((it, idx) => {
            const cv = it.column_values || [];
            const srcText = (cv.find(c => c.id === sourceId) || {}).text || '';
            const typText = (cv.find(c => c.id === typeId) || {}).text || '';
            const stsText = (cv.find(c => c.id === statusId) || {}).text || '';
            const created = new Date(it.created_at);

            Logger.log(`[${idx + 1}] #${it.id} "${it.name}"`);
            Logger.log(`    Source: "${srcText}" / Type: "${typText}" / Status: "${stsText}"`);
            Logger.log(`    Créé le: ${created.toISOString()}`);
        });

        Logger.log(`\n🎉 FILTRES MONDAY FONCTIONNENT AVEC LES INDEX !`);
        Logger.log(`💡 Prochaine étape: Intégrer dans SITE_mondayLeadCallsCountsByMonth_`);

        return { success: true, count: items.length, items };

    } catch (e) {
        Logger.log(`❌ ERREUR: ${e.message}`);
        Logger.log(e.stack);
        return { success: false, error: e.message };
    }
}
