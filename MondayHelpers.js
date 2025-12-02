/**
 * Helpers pour mapper les valeurs texte Monday vers les index
 * Nécessaire car les filtres Monday query_params utilisent des index, pas du texte
 */

/**
 * Récupère les settings d'une colonne Monday (labels + index)
 * @param {string} boardId - ID du board Monday
 * @param {string} columnId - ID de la colonne
 * @returns {Array} - Labels avec leur index [{id, index, name}, ...]
 */
function MONDAY_getColumnSettings_(boardId, columnId) {
    const q = `
    query($bid:[ID!]){
      boards(ids:$bid){
        columns(ids:["${columnId}"]){
          id
          title
          settings_str
        }
      }
    }`;

    const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)] });
    const col = (d && d.boards && d.boards[0] && d.boards[0].columns && d.boards[0].columns[0]) || null;

    if (!col || !col.settings_str) {
        Logger.log(`[Monday] Pas de settings pour colonne ${columnId}`);
        return [];
    }

    try {
        const settings = JSON.parse(col.settings_str);
        // Pour status: settings.labels = {0: {name: "X"}, 1: {name: "Y"}}
        // Pour color: settings.labels = {0: "X", 1: "Y", ...}
        const labels = settings.labels || {};

        const result = [];
        Object.keys(labels).forEach(key => {
            const idx = parseInt(key, 10);
            const label = labels[key];
            const name = typeof label === 'string' ? label : (label.name || '');
            result.push({ id: key, index: idx, name });
        });

        return result;
    } catch (e) {
        Logger.log(`[Monday] Erreur parse settings: ${e.message}`);
        return [];
    }
}

/**
 * Convertit une liste de textes en liste d'index Monday
 * @param {string} boardId - ID du board
 * @param {string} columnId - ID de la colonne
 * @param {Array<string>} textValues - Liste de textes à convertir
 * @returns {Array<number>} - Liste des index correspondants
 */
function MONDAY_textToIndexes_(boardId, columnId, textValues) {
    const labels = MONDAY_getColumnSettings_(boardId, columnId);
    const indexes = [];

    textValues.forEach(text => {
        const found = labels.find(l => l.name === text);
        if (found) {
            indexes.push(found.index);
        } else {
            Logger.log(`[Monday] ATTENTION: "${text}" non trouvé dans colonne ${columnId}`);
        }
    });

    return indexes;
}
