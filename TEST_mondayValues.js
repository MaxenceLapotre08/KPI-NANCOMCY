/**
 * TEST_mondayValues.js
 * Découvrir les VRAIES valeurs des colonnes Monday (index vs texte)
 */

function TEST_monday_inspect_values() {
    Logger.log('========== INSPECTION VALEURS MONDAY ==========');

    const boardId = SITE_MONDAY_LEADS_BOARD_ID;
    const sourceId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_SOURCE);
    const typeId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_TYPE);
    const statusId = SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_STATUS);

    Logger.log(`\nColonnes Monday:`);
    Logger.log(`- Source ID: ${sourceId}`);
    Logger.log(`- Type ID: ${typeId}`);
    Logger.log(`- Status ID: ${statusId}`);

    // Récupérer 20 items SANS FILTRE pour voir les valeurs réelles
    const q = `
    query($bid:[ID!]){
      boards(ids:$bid){
        items_page(limit:20){
          items{
            id name
            column_values(ids:["${sourceId}", "${typeId}", "${statusId}"]){ 
              id 
              text
              value
            }
          }
        }
      }
    }`;

    const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)] });
    const items = (d && d.boards && d.boards[0] && d.boards[0].items_page && d.boards[0].items_page.items) || [];

    Logger.log(`\n📦 ${items.length} items récupérés SANS filtre\n`);

    const uniqueValues = {
        source: new Set(),
        type: new Set(),
        status: new Set()
    };

    items.forEach((it, idx) => {
        const cv = it.column_values || [];
        const srcCol = cv.find(c => c.id === sourceId) || {};
        const typCol = cv.find(c => c.id === typeId) || {};
        const stsCol = cv.find(c => c.id === statusId) || {};

        Logger.log(`[${idx + 1}] Item: "${it.name}"`);
        Logger.log(`    Source - text: "${srcCol.text || ''}" | value: ${srcCol.value || 'null'}`);
        Logger.log(`    Type   - text: "${typCol.text || ''}" | value: ${typCol.value || 'null'}`);
        Logger.log(`    Status - text: "${stsCol.text || ''}" | value: ${stsCol.value || 'null'}`);

        if (srcCol.text) uniqueValues.source.add(srcCol.text);
        if (typCol.text) uniqueValues.type.add(typCol.text);
        if (stsCol.text) uniqueValues.status.add(stsCol.text);

        // Parser le value JSON pour voir les index
        if (srcCol.value) {
            try {
                const parsed = JSON.parse(srcCol.value);
                Logger.log(`    Source VALUE parsed: ${JSON.stringify(parsed)}`);
            } catch (e) { }
        }
        if (typCol.value) {
            try {
                const parsed = JSON.parse(typCol.value);
                Logger.log(`    Type VALUE parsed: ${JSON.stringify(parsed)}`);
            } catch (e) { }
        }
        if (stsCol.value) {
            try {
                const parsed = JSON.parse(stsCol.value);
                Logger.log(`    Status VALUE parsed: ${JSON.stringify(parsed)}`);
            } catch (e) { }
        }
        Logger.log('');
    });

    Logger.log('\n========== VALEURS UNIQUES TROUVÉES ==========');
    Logger.log(`Source: ${JSON.stringify(Array.from(uniqueValues.source))}`);
    Logger.log(`Type:   ${JSON.stringify(Array.from(uniqueValues.type))}`);
    Logger.log(`Status: ${JSON.stringify(Array.from(uniqueValues.status))}`);

    Logger.log('\n💡 IMPORTANT:');
    Logger.log('Pour les colonnes color/label, il faut peut-être utiliser les INDEX ({"label_ids":[X]})');
    Logger.log('Pour les colonnes status, il faut peut-être utiliser les INDEX ({"index":X})');

    return { uniqueValues, items };
}
