/**
 * RateLimiter.js
 * Gestion du rate limiting pour respecter les limites des APIs externes.
 * Utilise le pattern Token Bucket.
 */

const RateLimiter = {
    // Configuration des limites par API
    limits: {
        'matomo': { maxRequests: 10, intervalMs: 1000 },      // 10 req/s
        'monday': { maxRequests: 300, intervalMs: 60000 },    // 300 req/min
        'meta': { maxRequests: 200, intervalMs: 3600000 },    // 200 req/h
        'gads': { maxRequests: 2000, intervalMs: 60000 },     // 2000 req/min (approximatif)
        'magnetis': { maxRequests: 100, intervalMs: 60000 },  // 100 req/min
        'paperform': { maxRequests: 60, intervalMs: 60000 },  // 60 req/min
        'default': { maxRequests: 50, intervalMs: 60000 }     // 50 req/min
    },

    // Stockage des buckets (en mémoire pour la session)
    buckets: {},

    /**
     * Initialise un bucket pour une API.
     * @param {string} apiName - Nom de l'API
     */
    _initBucket(apiName) {
        const limit = this.limits[apiName] || this.limits.default;
        this.buckets[apiName] = {
            tokens: limit.maxRequests,
            maxTokens: limit.maxRequests,
            intervalMs: limit.intervalMs,
            lastRefill: Date.now()
        };
    },

    /**
     * Recharge les tokens du bucket.
     * @param {object} bucket - Bucket à recharger
     */
    _refillBucket(bucket) {
        const now = Date.now();
        const elapsed = now - bucket.lastRefill;

        if (elapsed >= bucket.intervalMs) {
            bucket.tokens = bucket.maxTokens;
            bucket.lastRefill = now;
            Logger.log(`[RateLimit] Bucket rechargé: ${bucket.tokens} tokens`);
        }
    },

    /**
     * Vérifie si une requête peut être effectuée.
     * @param {string} apiName - Nom de l'API
     * @returns {boolean} True si requête autorisée
     */
    canRequest(apiName) {
        if (!this.buckets[apiName]) {
            this._initBucket(apiName);
        }

        const bucket = this.buckets[apiName];
        this._refillBucket(bucket);

        return bucket.tokens > 0;
    },

    /**
     * Consomme un token pour une requête.
     * @param {string} apiName - Nom de l'API
     * @returns {boolean} True si token consommé
     */
    consume(apiName) {
        if (!this.buckets[apiName]) {
            this._initBucket(apiName);
        }

        const bucket = this.buckets[apiName];
        this._refillBucket(bucket);

        if (bucket.tokens > 0) {
            bucket.tokens--;
            Logger.log(`[RateLimit] ${apiName}: ${bucket.tokens}/${bucket.maxTokens} tokens restants`);
            return true;
        }

        return false;
    },

    /**
     * Attend si nécessaire avant d'autoriser une requête.
     * @param {string} apiName - Nom de l'API
     */
    waitIfNeeded(apiName) {
        const maxWait = 60000; // Max 60s d'attente
        const checkInterval = 100; // Vérifier tous les 100ms
        const startWait = Date.now();

        while (!this.canRequest(apiName)) {
            const elapsed = Date.now() - startWait;
            if (elapsed > maxWait) {
                throw new Error(`[RateLimit] Timeout après ${maxWait}ms pour ${apiName}`);
            }

            Logger.log(`[RateLimit] Attente pour ${apiName}... (${elapsed}ms)`);
            Utilities.sleep(checkInterval);
        }

        this.consume(apiName);
    },

    /**
     * Wrapper pour appliquer le rate limiting automatiquement.
     * @param {string} apiName - Nom de l'API
     * @param {Function} fn - Fonction à exécuter
     * @returns {*} Résultat de la fonction
     */
    throttle(apiName, fn) {
        this.waitIfNeeded(apiName);
        return fn();
    },

    /**
     * Réinitialise tous les buckets.
     */
    reset() {
        this.buckets = {};
        Logger.log('[RateLimit] Tous les buckets réinitialisés');
    }
};

// Export pour compatibilité globale
if (typeof module !== 'undefined' && module.exports) {
    module.exports = RateLimiter;
}
