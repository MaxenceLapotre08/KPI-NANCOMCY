/**
 * Validators.js
 * Validation des paramètres d'entrée pour garantir la robustesse du code.
 * Évite les erreurs silencieuses et facilite le debugging.
 */

const Validators = {
    /**
     * Valide une plage de dates.
     * @param {Date} startDate - Date de début
     * @param {Date} endDate - Date de fin
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si les dates sont invalides ou incohérentes
     * @returns {boolean} True si valide
     */
    validateDateRange(startDate, endDate, context = '') {
        const prefix = context ? `[${context}] ` : '';

        // Vérifier que ce sont des objets Date
        if (!(startDate instanceof Date)) {
            throw new Error(`${prefix}startDate doit être un objet Date (reçu: ${typeof startDate})`);
        }

        if (!(endDate instanceof Date)) {
            throw new Error(`${prefix}endDate doit être un objet Date (reçu: ${typeof endDate})`);
        }

        // Vérifier que les dates sont valides
        if (isNaN(startDate.getTime())) {
            throw new Error(`${prefix}startDate est une date invalide`);
        }

        if (isNaN(endDate.getTime())) {
            throw new Error(`${prefix}endDate est une date invalide`);
        }

        // Vérifier que startDate <= endDate
        if (startDate > endDate) {
            const startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            const endStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            throw new Error(`${prefix}startDate (${startStr}) doit être antérieure ou égale à endDate (${endStr})`);
        }

        return true;
    },

    /**
     * Valide qu'une propriété de script existe.
     * @param {string} propertyName - Nom de la propriété
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si la propriété n'existe pas
     * @returns {string} Valeur de la propriété
     */
    validateProperty(propertyName, context = '') {
        const prefix = context ? `[${context}] ` : '';
        const value = PropertiesService.getScriptProperties().getProperty(propertyName);

        if (!value || value.trim() === '') {
            throw new Error(
                `${prefix}Propriété de script manquante: ${propertyName}. ` +
                `Veuillez la configurer dans Projet > Paramètres > Propriétés du script.`
            );
        }

        return value;
    },

    /**
     * Valide qu'une valeur n'est ni null ni undefined.
     * @param {*} value - Valeur à vérifier
     * @param {string} fieldName - Nom du champ pour le message d'erreur
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si la valeur est null ou undefined
     * @returns {*} La valeur si elle est valide
     */
    validateRequired(value, fieldName, context = '') {
        const prefix = context ? `[${context}] ` : '';

        if (value === null || value === undefined) {
            throw new Error(`${prefix}${fieldName} est requis (reçu: ${value})`);
        }

        return value;
    },

    /**
     * Valide qu'une chaîne n'est pas vide.
     * @param {string} value - Chaîne à vérifier
     * @param {string} fieldName - Nom du champ pour le message d'erreur
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si la chaîne est vide
     * @returns {string} La chaîne si elle est valide
     */
    validateNonEmpty(value, fieldName, context = '') {
        const prefix = context ? `[${context}] ` : '';

        if (typeof value !== 'string') {
            throw new Error(`${prefix}${fieldName} doit être une chaîne (reçu: ${typeof value})`);
        }

        if (value.trim() === '') {
            throw new Error(`${prefix}${fieldName} ne peut pas être vide`);
        }

        return value;
    },

    /**
     * Valide qu'un nombre est dans une plage.
     * @param {number} value - Nombre à vérifier
     * @param {number} min - Valeur minimale (inclusive)
     * @param {number} max - Valeur maximale (inclusive)
     * @param {string} fieldName - Nom du champ pour le message d'erreur
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si le nombre est hors de la plage
     * @returns {number} Le nombre si valide
     */
    validateRange(value, min, max, fieldName, context = '') {
        const prefix = context ? `[${context}] ` : '';

        if (typeof value !== 'number' || isNaN(value)) {
            throw new Error(`${prefix}${fieldName} doit être un nombre (reçu: ${value})`);
        }

        if (value < min || value > max) {
            throw new Error(`${prefix}${fieldName} doit être entre ${min} et ${max} (reçu: ${value})`);
        }

        return value;
    },

    /**
     * Valide qu'une feuille Google Sheets existe.
     * @param {Spreadsheet} spreadsheet - Document Spreadsheet
     * @param {string} sheetName - Nom de la feuille
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si la feuille n'existe pas
     * @returns {Sheet} La feuille si elle existe
     */
    validateSheetExists(spreadsheet, sheetName, context = '') {
        const prefix = context ? `[${context}] ` : '';
        const sheet = spreadsheet.getSheetByName(sheetName);

        if (!sheet) {
            const availableSheets = spreadsheet.getSheets().map(s => s.getName()).join(', ');
            throw new Error(
                `${prefix}Feuille "${sheetName}" introuvable. ` +
                `Feuilles disponibles: ${availableSheets}`
            );
        }

        return sheet;
    },

    /**
     * Valide un format de mois YYYY-MM.
     * @param {string} monthKey - Clé du mois
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si le format est invalide
     * @returns {string} La clé si valide
     */
    validateMonthKey(monthKey, context = '') {
        const prefix = context ? `[${context}] ` : '';

        if (typeof monthKey !== 'string') {
            throw new Error(`${prefix}monthKey doit être une chaîne (reçu: ${typeof monthKey})`);
        }

        const regex = /^\d{4}-\d{2}$/;
        if (!regex.test(monthKey)) {
            throw new Error(`${prefix}Format monthKey invalide: "${monthKey}" (attendu: YYYY-MM)`);
        }

        // Vérifier que le mois est entre 01 et 12
        const month = parseInt(monthKey.slice(5, 7), 10);
        if (month < 1 || month > 12) {
            throw new Error(`${prefix}Mois invalide dans "${monthKey}" (doit être entre 01 et 12)`);
        }

        return monthKey;
    },

    /**
     * Valide qu'un tableau n'est pas vide.
     * @param {Array} array - Tableau à vérifier
     * @param {string} fieldName - Nom du champ pour le message d'erreur
     * @param {string} context - Contexte pour le message d'erreur
     * @throws {Error} Si le tableau est vide ou n'est pas un tableau
     * @returns {Array} Le tableau si valide
     */
    validateNonEmptyArray(array, fieldName, context = '') {
        const prefix = context ? `[${context}] ` : '';

        if (!Array.isArray(array)) {
            throw new Error(`${prefix}${fieldName} doit être un tableau (reçu: ${typeof array})`);
        }

        if (array.length === 0) {
            throw new Error(`${prefix}${fieldName} ne peut pas être vide`);
        }

        return array;
    }
};

// Export pour compatibilité globale
if (typeof module !== 'undefined' && module.exports) {
    module.exports = Validators;
}
