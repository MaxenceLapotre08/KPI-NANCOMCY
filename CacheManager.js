/**
 * CacheManager.js
 * Gestion centralisée du cache pour optimiser les performances.
 * Utilise CacheService de Google Apps Script.
 */

const CacheManager = {
    // Durées de cache en secondes
    CACHE_DURATION: {
        SHORT: 300,      // 5 minutes
        MEDIUM: 1800,    // 30 minutes
        LONG: 21600      // 6 heures (max Apps Script)
    },

    /**
     * Récupère une valeur du cache.
     * @param {string} key - Clé du cache
     * @returns {*} Valeur cachée ou null si non trouvée/expirée
     */
    get(key) {
        try {
            const cache = CacheService.getScriptCache();
            const cached = cache.get(key);

            if (cached) {
                Logger.log(`[Cache HIT] ${key}`);
                return JSON.parse(cached);
            }

            Logger.log(`[Cache MISS] ${key}`);
            return null;
        } catch (error) {
            ErrorHandler.logWarning('CacheManager.get', `Erreur lecture cache: ${error.message}`, { key });
            return null;
        }
    },

    /**
     * Stocke une valeur dans le cache.
     * @param {string} key - Clé du cache
     * @param {*} value - Valeur à cacher
     * @param {number} duration - Durée en secondes (défaut: MEDIUM)
     * @returns {boolean} True si succès
     */
    set(key, value, duration = null) {
        try {
            const cache = CacheService.getScriptCache();
            const ttl = duration || this.CACHE_DURATION.MEDIUM;

            cache.put(key, JSON.stringify(value), ttl);
            Logger.log(`[Cache SET] ${key} (TTL: ${ttl}s)`);
            return true;
        } catch (error) {
            ErrorHandler.logWarning('CacheManager.set', `Erreur écriture cache: ${error.message}`, { key });
            return false;
        }
    },

    /**
     * Invalide une clé du cache.
     * @param {string} key - Clé à invalider
     */
    invalidate(key) {
        try {
            const cache = CacheService.getScriptCache();
            cache.remove(key);
            Logger.log(`[Cache INVALIDATE] ${key}`);
        } catch (error) {
            ErrorHandler.logWarning('CacheManager.invalidate', `Erreur invalidation cache: ${error.message}`, { key });
        }
    },

    /**
     * Vide tout le cache.
     */
    clear() {
        try {
            const cache = CacheService.getScriptCache();
            cache.removeAll();
            Logger.log('[Cache] Cache vidé complètement');
        } catch (error) {
            ErrorHandler.logError('CacheManager.clear', error, 'ERROR');
        }
    },

    /**
     * Wrapper pour fetch avec cache automatique.
     * @param {string} cacheKey - Clé unique pour cette requête
     * @param {Function} fetchFn - Fonction qui récupère les données
     * @param {number} duration - Durée du cache
     * @returns {*} Données (du cache ou fraîches)
     */
    cachedFetch(cacheKey, fetchFn, duration = null) {
        // 1. Essayer le cache
        const cached = this.get(cacheKey);
        if (cached !== null) {
            return cached;
        }

        // 2. Récupérer les données
        const data = fetchFn();

        // 3. Mettre en cache
        this.set(cacheKey, data, duration);

        return data;
    },

    /**
     * Cache pour les colonnes Monday (évite appels répétés à l'API).
     * @param {number} boardId - ID du board Monday
     * @returns {Array} Colonnes du board
     */
    getMondayBoardColumns(boardId) {
        const cacheKey = `monday_columns_${boardId}`;

        return this.cachedFetch(
            cacheKey,
            () => {
                // GraphQL pour récupérer les colonnes
                const query = `query($bid:[ID!]){ boards(ids:$bid){ columns{ id title type } } }`;
                const data = Utils_mondayGraphQL(query, { bid: [Number(boardId)] });
                return (data && data.boards && data.boards[0] && data.boards[0].columns) || [];
            },
            this.CACHE_DURATION.LONG // 6h - les colonnes changent rarement
        );
    },

    /**
     * Cache pour les reviews GMB (évite de récupérer toutes les reviews à chaque fois).
     * @param {string} locationId - ID de la location GMB
     * @param {Date} since - Date depuis laquelle récupérer
     * @returns {Array} Reviews
     */
    getGMBReviews(locationId, since) {
        const cacheKey = `gmb_reviews_${locationId}_${since.toISOString().slice(0, 10)}`;

        return this.cachedFetch(
            cacheKey,
            () => {
                // Fonction de récupération GMB (à implémenter dans GMB.js)
                if (typeof GMB_fetchReviews_ === 'function') {
                    return GMB_fetchReviews_(locationId, since);
                }
                return [];
            },
            this.CACHE_DURATION.MEDIUM // 30min - les reviews peuvent arriver
        );
    }
};

// Export pour compatibilité globale
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CacheManager;
}
