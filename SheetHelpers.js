/**
 * SheetHelpers.js
 * Fonctions réutilisables pour la manipulation de Google Sheets.
 * Évite la duplication de code entre Site Internet.js, Google Ads.js, Meta Ads.js et GMB.js.
 */

const SheetHelpers = {
    /**
     * Style une ligne année (fond coloré et texte en gras).
     * @param {Sheet} sheet - Feuille Google Sheets
     * @param {number} row - Numéro de ligne (1-indexed)
     * @param {number} moisCol - Colonne du mois (1-indexed)
     */
    styleYearRow(sheet, row, moisCol) {
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e6e1f5');
        sheet.getRange(row, moisCol).setFontWeight('bold');
    },

    /**
     * Trouve une ligne existante pour un mois donné (sans insertion).
     * @param {Sheet} sheet - Feuille Google Sheets
     * @param {number} moisCol - Colonne du mois (1-indexed)
     * @param {string} targetYM - Mois au format "YYYY-MM"
     * @param {number} startRow - Première ligne de données
     * @returns {number} Numéro de ligne (0 si non trouvé)
     */
    findExistingMonthRow(sheet, moisCol, targetYM, startRow = 6) {
        const lastRow = sheet.getLastRow();
        if (lastRow < startRow) return 0;

        for (let r = startRow; r <= lastRow; r++) {
            const cellValue = sheet.getRange(r, moisCol).getValue();
            if (Utils_isYearSeparatorRow(cellValue)) continue;

            const ym = Utils_sheetCellToYYYYMM(cellValue);
            if (ym === targetYM) return r;
        }

        return 0;
    },

    /**
     * Garantit qu'une ligne existe pour un mois donné.
     * Crée la ligne année si nécessaire (2024, 2025...) et insère le mois.
     * @param {Sheet} sheet - Feuille Google Sheets
     * @param {number} moisCol - Colonne du mois (1-indexed)
     * @param {string} targetYM - Mois au format "YYYY-MM"
     * @param {number} startRow - Première ligne de données
     * @param {string} logPrefix - Préfixe pour les logs (ex: "Ads", "Meta")
     * @returns {number} Numéro de ligne (1-indexed)
     */
    ensureMonthRow(sheet, moisCol, targetYM, startRow = 6, logPrefix = 'Sheet') {
        const lastRow = sheet.getLastRow();

        // 1. Chercher si la ligne du mois existe déjà
        for (let r = startRow; r <= lastRow; r++) {
            const cellValue = sheet.getRange(r, moisCol).getValue();
            if (Utils_isYearSeparatorRow(cellValue)) continue;

            const ym = Utils_sheetCellToYYYYMM(cellValue);
            if (ym === targetYM) {
                Logger.log(`[${logPrefix}/INFO] Ligne pour ${targetYM} trouvée en ${r}. Utilisation de la ligne existante.`);
                return r;
            }
        }

        // 2. Trouver la dernière ligne de données
        let lastDataRow = startRow - 1;
        for (let r = lastRow; r >= startRow; r--) {
            const cellValue = sheet.getRange(r, moisCol).getValue();
            if (cellValue !== '') {
                lastDataRow = r;
                break;
            }
        }

        const targetRow = lastDataRow + 1;
        const lastDataValue = lastDataRow >= startRow ? sheet.getRange(lastDataRow, moisCol).getValue() : '';
        const lastDataYM = Utils_sheetCellToYYYYMM(lastDataValue) || '1999-12';

        // 3. Gérer le changement d'année
        const lastYear = parseInt(lastDataYM.slice(0, 4), 10);
        const targetYear = parseInt(targetYM.slice(0, 4), 10);

        if (targetYear > lastYear) {
            Logger.log(`[${logPrefix}/INFO] Changement d'année détecté (${lastYear} -> ${targetYear}). Ajout d'un séparateur.`);
            sheet.getRange(targetRow, moisCol).setValue(String(targetYear));
            this.styleYearRow(sheet, targetRow, moisCol);
            Logger.log(`[${logPrefix}/WRITE] Écriture des KPIs pour ${targetYM} sur la nouvelle ligne ${targetRow + 1}`);
            return targetRow + 1;
        }

        // 4. Cas normal : écriture sur la ligne suivante
        Logger.log(`[${logPrefix}/WRITE] Écriture des KPIs pour ${targetYM} sur la nouvelle ligne ${targetRow}`);
        return targetRow;
    },

    /**
     * Vérifie si une colonne est protégée (contient des formules à préserver).
     * @param {Sheet} sheet - Feuille Google Sheets
     * @param {number} col - Numéro de colonne (1-indexed)
     * @param {string[]} protectedHeaders - Liste des en-têtes protégés (normalisés)
     * @param {number} headerRow - Ligne d'en-têtes
     * @returns {boolean} True si la colonne est protégée
     */
    isProtectedColumn(sheet, col, protectedHeaders, headerRow = 3) {
        if (!col) return false;

        const headerValue = sheet.getRange(headerRow, col).getValue();
        const normalized = Utils_normHeader(headerValue);

        return protectedHeaders.includes(normalized);
    },

    /**
     * Écrit une valeur dans une cellule uniquement si elle ne contient pas de formule.
     * Wrapper pour Utils_setPreserveFormula avec validation et logging.
     * @param {Sheet} sheet - Feuille Google Sheets
     * @param {number} row - Numéro de ligne (1-indexed)
     * @param {number} col - Numéro de colonne (1-indexed)
     * @param {*} value - Valeur à écrire
     * @param {string} numberFormat - Format de nombre optionnel
     * @param {boolean} skipIfProtected - Si true, saute l'écriture si protégée
     * @returns {boolean} True si la valeur a été écrite
     */
    writeCell(sheet, row, col, value, numberFormat = null, skipIfProtected = false) {
        if (!col || col === 0) {
            Logger.log(`[SheetHelpers/WARNING] Colonne invalide: ${col}`);
            return false;
        }

        try {
            Utils_setPreserveFormula(sheet, row, col, value, numberFormat);
            return true;
        } catch (error) {
            Logger.log(`[SheetHelpers/ERROR] Échec écriture cellule [${row},${col}]: ${error.message}`);
            return false;
        }
    },

    /**
     * Génère une liste de clés de mois entre deux dates.
     * @param {Date} fromDate - Date de début
     * @param {Date} toDate - Date de fin
     * @returns {string[]} Liste de mois au format "YYYY-MM"
     */
    monthRangeKeys(fromDate, toDate) {
        const keys = [];
        const current = new Date(fromDate.getFullYear(), fromDate.getMonth(), 1);
        const end = new Date(toDate.getFullYear(), toDate.getMonth(), 1);

        while (current <= end) {
            const ym = `${current.getFullYear()}-${String(current.getMonth() + 1).padStart(2, '0')}`;
            keys.push(ym);
            current.setMonth(current.getMonth() + 1);
        }

        return keys;
    },

    /**
     * Trouve une colonne de manière sécurisée en évitant les colonnes interdites.
     * Utile pour "Budget investi" qui peut être confondu avec d'autres colonnes.
     * @param {Sheet} sheet - Feuille Google Sheets
     * @param {string[]} targetHeaders - En-têtes recherchés
     * @param {string[]} forbiddenHeaders - En-têtes à éviter
     * @param {number} headerRow - Ligne d'en-têtes
     * @returns {number} Index de colonne (1-indexed) ou 0 si non trouvé
     */
    findColumnSafe(sheet, targetHeaders, forbiddenHeaders = [], headerRow = 3) {
        const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
        const normalizedHeaders = headers.map(Utils_normHeader);

        // Normaliser les listes
        const targets = targetHeaders.map(Utils_normHeader);
        const forbidden = forbiddenHeaders.map(Utils_normHeader);

        // 1. Chercher match exact en évitant les interdits
        for (let i = 0; i < normalizedHeaders.length; i++) {
            const h = normalizedHeaders[i];
            if (!h || forbidden.some(f => h.includes(f))) continue;
            if (targets.includes(h)) return i + 1;
        }

        // 2. Chercher match partiel en évitant les interdits
        for (let i = 0; i < normalizedHeaders.length; i++) {
            const h = normalizedHeaders[i];
            if (!h || forbidden.some(f => h.includes(f))) continue;
            if (targets.some(t => h.includes(t))) return i + 1;
        }

        return 0;
    }
};

// Export pour compatibilité globale
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SheetHelpers;
}
