/**
 * Config.js
 * Configuration centralisée pour tous les scripts du projet KPI NANCOMCY.
 * Évite la duplication de constantes et facilite la maintenance.
 */

const CONFIG = {
    // ========== Timeouts et retries ==========
    DEFAULT_RETRY_COUNT: 3,
    DEFAULT_RETRY_DELAY_MS: 1000,
    API_TIMEOUT_MS: 30000,

    // ========== Pagination ==========
    DEFAULT_PAGE_SIZE: 100,
    MAX_PAGE_SIZE: 500,
    PAPERFORM_PAGE_SIZE: 100,
    MONDAY_PAGE_SIZE: 500,
    MAGNETIS_PAGE_SIZE: 250,

    // ========== Dates ==========
    DEFAULT_MONTHS_BACK: 21,
    MATOMO_CUTOFF_DATE: new Date(2025, 10, 1), // 1er Novembre 2025

    // ========== Structure des feuilles ==========
    DEFAULT_HEADER_ROW: 3,
    DEFAULT_START_ROW: 6,

    // ========== URLs d'API ==========
    MONDAY_API_URL: 'https://api.monday.com/v2',
    MONDAY_API_VERSION: '2023-10',
    MATOMO_BASE_URL: 'https://matomo.aleo.agency',
    MAGNETIS_BASE_URL: 'https://api.magnetis.io',
    PAPERFORM_BASE_URL: 'https://api.paperform.co/v1',

    // ========== IDs Matomo ==========
    MATOMO_SITE_ID: 1492,

    // ========== Logging ==========
    ENABLE_VERBOSE_LOGGING: false,
    LOG_API_REQUESTS: true,
    LOG_API_RESPONSES: false,

    // ========== Feuilles Google Sheets ==========
    SHEET_NAME_SITE: 'Site Internet',
    SHEET_NAME_ADS: 'Google Ads',
    SHEET_NAME_META: 'Meta Ads',
    SHEET_NAME_GMB: 'GMB'
};

/**
 * Récupère une valeur de configuration.
 * @param {string} key - Clé de configuration
 * @returns {*} Valeur de configuration
 */
function getConfig(key) {
    if (CONFIG.hasOwnProperty(key)) {
        return CONFIG[key];
    }
    throw new Error(`Configuration key '${key}' not found in CONFIG`);
}
