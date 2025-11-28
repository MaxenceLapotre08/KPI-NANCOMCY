/**
 * DEBUG_mondayLeadCalls.js
 * Fonction de debug pour identifier pourquoi les appels lead ne remontent pas.
 */

function DEBUG_mondayLeadCalls() {
    const boardId = SITE_MONDAY_LEADS_BOARD_ID;
    const colSource = SITE_MONDAY_LEADS_COL_SOURCE;
    const colType = SITE_MONDAY_LEADS_COL_TYPE;
    const colStatus = SITE_MONDAY_LEADS_COL_STATUS;

    const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
    const typeId = SITE_mondayResolveColumnId_(boardId, colType);
    const statusId = SITE_mondayResolveColumnId_(boardId, colStatus);
    const dateId = SITE_MONDAY_LEADS_COL_DATE ? SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_DATE) : null;

    Logger.log(`============ DEBUG MONDAY LEAD CALLS ============`);
    Logger.log(`Board ID: ${boardId}`);
    Logger.log(`Source Column: "${colSource}" -> ID: ${sourceId}`);
    Logger.log(`Type Column: "${colType}" -> ID: ${typeId}`);
    Logger.log(`Status Column: "${colStatus}" -> ID: ${statusId}`);
    Logger.log(`Date Column: "${SITE_MONDAY_LEADS_COL_DATE}" -> ID: ${dateId}`);
    Logger.log(`\nCritères de filtrage:`);
    Logger.log(`- Source MATCH: ${JSON.stringify(SITE_MONDAY_LEADS_SOURCE_MATCH)}`);
    Logger.log(`- Type MATCH: ${JSON.stringify(SITE_MONDAY_LEADS_TYPE_MATCH)}`);
    Logger.log(`- Status MATCH: ${JSON.stringify(SITE_MONDAY_LEADS_STATUS_MATCH)}`);

    // Période de test: mois en cours
    const today = new Date();
    const fromUTC = new Date(Date.UTC(today.getFullYear(), today.getMonth(), 1));
    const toUTC = new Date(Date.UTC(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59));

    Logger.log(`\nPériode: ${fromUTC.toISOString()} -> ${toUTC.toISOString()}`);

    let totalItems = 0;
    let matchedItems = 0;
    let cursor = null;

    const q = `
    query($bid:[ID!], $cursor:String, $cols:[String!]){
      boards(ids:$bid){
        items_page(limit:50, cursor:$cursor){
          cursor
          items{
            id name state created_at
            column_values(ids:$cols){ id text value }
          }
        }
      }
    }`;

    const colIds = [sourceId, typeId].concat(statusId ? [statusId] : []).concat(dateId ? [dateId] : []);

    do {
        const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)], cursor, cols: colIds });
        const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
        if (!page) break;

        (page.items || []).forEach(it => {
            totalItems++;

            if (String(it.state || '').toLowerCase() === 'archived') {
                Logger.log(`\n[${totalItems}] Item #${it.id} "${it.name}" - ARCHIVÉ (ignoré)`);
                return;
            }

            const createdAt = new Date(it.created_at);
            const cv = it.column_values || [];
            const srcText = (cv.find(c => c.id === sourceId) || {}).text || '';
            const typText = (cv.find(c => c.id === typeId) || {}).text || '';
            const sts = statusId ? ((cv.find(c => c.id === statusId) || {}).text || '') : '';
            const dcv = dateId ? ((cv.find(c => c.id === dateId) || {}).value || '') : '';

            // Essayer de parser la date Monday si disponible
            let when = createdAt;
            if (dateId && dcv) {
                try {
                    // Format Monday: {"date":"2025-11-28","time":"12:30:00"}
                    const parsed = JSON.parse(dcv);
                    if (parsed && parsed.date) {
                        when = new Date(parsed.date);
                    }
                } catch (e) {
                    // Si parsing échoue, utiliser created_at
                }
            }

            Logger.log(`\n[${totalItems}] Item #${it.id} "${it.name}"`);
            Logger.log(`  Source: "${srcText}"`);
            Logger.log(`  Type: "${typText}"`);
            Logger.log(`  Statut: "${sts}"`);
            Logger.log(`  Date colonne: ${dcv || 'N/A'}`);
            Logger.log(`  Date utilisée: ${when.toISOString()}`);
            Logger.log(`  Created at: ${createdAt.toISOString()}`);

            // Vérifier les critères
            const sourceMatch = Utils_textIncludesAny(srcText, SITE_MONDAY_LEADS_SOURCE_MATCH);
            const typeMatch = Utils_textIncludesAny(typText, SITE_MONDAY_LEADS_TYPE_MATCH);
            const statusMatch = statusId ? Utils_textIncludesAny(sts, SITE_MONDAY_LEADS_STATUS_MATCH) : true;
            const dateMatch = when >= fromUTC && when <= toUTC;

            Logger.log(`  ✓ Source match: ${sourceMatch}`);
            Logger.log(`  ✓ Type match: ${typeMatch}`);
            Logger.log(`  ✓ Status match: ${statusMatch}`);
            Logger.log(`  ✓ Date match: ${dateMatch}`);

            if (sourceMatch && typeMatch && statusMatch && dateMatch) {
                matchedItems++;
                Logger.log(`  ✅ COMPTABILISÉ !`);
            } else {
                Logger.log(`  ❌ NON COMPTABILISÉ`);
            }
        });

        cursor = page.cursor;
        if (totalItems >= 50) break; // Limiter à 50 items pour le debug
    } while (cursor);

    Logger.log(`\n============ RÉSUMÉ ============`);
    Logger.log(`Total items analysés: ${totalItems}`);
    Logger.log(`Items matchés (appels lead): ${matchedItems}`);
    Logger.log(`================================`);

    return { totalItems, matchedItems };
}
