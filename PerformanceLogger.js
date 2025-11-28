/**
 * PerformanceLogger.js
 * Logger de performance pour identifier les goulots d'étranglement.
 * Mesure le temps d'exécution des fonctions critiques.
 */

const PerformanceLogger = {
    // Stockage des mesures
    metrics: {},

    /**
     * Démarre une mesure de performance.
     * @param {string} label - Nom de la mesure
     * @returns {number} Timestamp de début
     */
    start(label) {
        const startTime = Date.now();
        this.metrics[label] = {
            start: startTime,
            end: null,
            duration: null
        };
        Logger.log(`[PERF START] ${label}`);
        return startTime;
    },

    /**
     * Termine une mesure de performance.
     * @param {string} label - Nom de la mesure
     * @returns {number} Durée en ms
     */
    end(label) {
        const endTime = Date.now();

        if (!this.metrics[label]) {
            Logger.log(`[PERF WARNING] Mesure '${label}' non démarrée`);
            return 0;
        }

        const metric = this.metrics[label];
        metric.end = endTime;
        metric.duration = endTime - metric.start;

        const duration = metric.duration;
        const seconds = (duration / 1000).toFixed(2);

        // Log avec couleur selon durée
        if (duration > 5000) {
            Logger.log(`[PERF END] ⚠️ ${label}: ${seconds}s (LENT)`);
        } else if (duration > 1000) {
            Logger.log(`[PERF END] ⚡ ${label}: ${seconds}s`);
        } else {
            Logger.log(`[PERF END] ✅ ${label}: ${duration}ms`);
        }

        return duration;
    },

    /**
     * Wrapper pour mesurer automatiquement une fonction.
     * @param {string} label - Nom de la mesure
     * @param {Function} fn - Fonction à mesurer
     * @returns {*} Résultat de la fonction
     */
    measure(label, fn) {
        this.start(label);
        try {
            const result = fn();
            this.end(label);
            return result;
        } catch (error) {
            this.end(label);
            throw error;
        }
    },

    /**
     * Récupère toutes les métriques enregistrées.
     * @returns {object} Métriques
     */
    getMetrics() {
        return this.metrics;
    },

    /**
     * Génère un rapport de performance.
     * @returns {string} Rapport formaté
     */
    report() {
        const lines = ['========================================'];
        lines.push('RAPPORT DE PERFORMANCE');
        lines.push('========================================');

        const sorted = Object.entries(this.metrics)
            .filter(([_, m]) => m.duration !== null)
            .sort((a, b) => b[1].duration - a[1].duration);

        if (sorted.length === 0) {
            lines.push('Aucune mesure enregistrée');
        } else {
            sorted.forEach(([label, metric]) => {
                const seconds = (metric.duration / 1000).toFixed(2);
                const status = metric.duration > 5000 ? '⚠️' : metric.duration > 1000 ? '⚡' : '✅';
                lines.push(`${status} ${label}: ${seconds}s`);
            });

            const total = sorted.reduce((sum, [_, m]) => sum + m.duration, 0);
            const totalSeconds = (total / 1000).toFixed(2);
            lines.push('----------------------------------------');
            lines.push(`TOTAL: ${totalSeconds}s`);
        }

        lines.push('========================================');
        const report = lines.join('\n');
        Logger.log(report);
        return report;
    },

    /**
     * Réinitialise toutes les métriques.
     */
    reset() {
        this.metrics = {};
        Logger.log('[PERF] Métriques réinitialisées');
    }
};

// Export pour compatibilité globale
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PerformanceLogger;
}
