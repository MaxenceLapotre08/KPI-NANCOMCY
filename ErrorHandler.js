/**
 * ErrorHandler.js
 * Gestion centralisée des erreurs avec logging structuré et alerting par email.
 * Permet de tracer toutes les erreurs du projet avec leur contexte.
 */

const ErrorHandler = {
    /**
     * Log une erreur de manière structurée avec contexte.
     * @param {string} context - Contexte de l'erreur (nom de la fonction, module, etc.)
     * @param {Error|string} error - L'erreur à logger
     * @param {string} severity - Niveau de gravité: 'INFO', 'WARNING', 'ERROR', 'CRITICAL'
     * @param {object} metadata - Métadonnées additionnelles (paramètres, état, etc.)
     */
    logError(context, error, severity = 'ERROR', metadata = {}) {
        const timestamp = new Date().toISOString();
        const errorMessage = error instanceof Error ? error.message : String(error);
        const stackTrace = error instanceof Error ? error.stack : '';

        const logEntry = {
            timestamp,
            context,
            severity,
            message: errorMessage,
            stack: stackTrace,
            metadata
        };

        // Log structuré pour Apps Script
        const logLine = `[${severity}] [${timestamp}] [${context}] ${errorMessage}`;
        Logger.log(logLine);

        if (stackTrace) {
            Logger.log(`Stack trace: ${stackTrace}`);
        }

        if (Object.keys(metadata).length > 0) {
            Logger.log(`Metadata: ${JSON.stringify(metadata, null, 2)}`);
        }

        // Si c'est une erreur critique, envoyer une alerte email
        if (severity === 'CRITICAL') {
            try {
                this.sendAlert(
                    `🚨 Erreur Critique - ${context}`,
                    `Une erreur critique s'est produite dans le projet KPI NANCOMCY.\n\n` +
                    `Contexte: ${context}\n` +
                    `Message: ${errorMessage}\n` +
                    `Heure: ${timestamp}\n\n` +
                    `Stack trace:\n${stackTrace}\n\n` +
                    `Métadonnées:\n${JSON.stringify(metadata, null, 2)}`
                );
            } catch (emailError) {
                Logger.log(`[ERROR] Impossible d'envoyer l'email d'alerte: ${emailError}`);
            }
        }

        return logEntry;
    },

    /**
     * Envoie une alerte par email.
     * @param {string} subject - Sujet de l'email
     * @param {string} message - Corps du message
     * @param {string[]} recipients - Liste d'emails (optionnel, sinon utilise l'email du propriétaire)
     */
    sendAlert(subject, message, recipients = null) {
        try {
            const emailRecipients = recipients || [Session.getActiveUser().getEmail()];

            if (!emailRecipients || emailRecipients.length === 0) {
                Logger.log('[WARNING] Aucun destinataire pour l\'email d\'alerte');
                return false;
            }

            MailApp.sendEmail({
                to: emailRecipients.join(','),
                subject: `[KPI NANCOMCY] ${subject}`,
                body: message,
                noReply: true
            });

            Logger.log(`[INFO] Email d'alerte envoyé à: ${emailRecipients.join(', ')}`);
            return true;
        } catch (error) {
            Logger.log(`[ERROR] Échec envoi email: ${error.message}`);
            return false;
        }
    },

    /**
     * Wrapper pour exécuter une fonction avec gestion d'erreurs automatique.
     * @param {Function} fn - Fonction à exécuter
     * @param {string} context - Contexte d'exécution
     * @param {object} options - Options: { severity, metadata, rethrow }
     * @returns {*} Résultat de la fonction ou null si erreur
     */
    wrapFunction(fn, context, options = {}) {
        const {
            severity = 'ERROR',
            metadata = {},
            rethrow = true
        } = options;

        try {
            return fn();
        } catch (error) {
            this.logError(context, error, severity, metadata);

            if (rethrow) {
                throw error;
            }

            return null;
        }
    },

    /**
     * Wrapper async/retry pour fonctions avec gestion d'erreurs et retry automatique.
     * Compatible avec Utils_executeWithRetry mais avec logging amélioré.
     * @param {Function} fn - Fonction à exécuter
     * @param {string} context - Contexte d'exécution
     * @param {number} maxRetries - Nombre max de tentatives
     * @param {number} initialDelay - Délai initial en ms
     * @returns {*} Résultat de la fonction
     */
    executeWithRetry(fn, context, maxRetries = 3, initialDelay = 1000) {
        let delay = initialDelay;
        let lastError = null;

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                Logger.log(`[${context}] Tentative ${attempt}/${maxRetries}`);
                const result = fn();

                if (attempt > 1) {
                    Logger.log(`[${context}] ✅ Succès après ${attempt} tentative(s)`);
                }

                return result;
            } catch (error) {
                lastError = error;

                this.logError(
                    context,
                    error,
                    attempt < maxRetries ? 'WARNING' : 'ERROR',
                    {
                        attempt,
                        maxRetries,
                        willRetry: attempt < maxRetries
                    }
                );

                if (attempt < maxRetries) {
                    Logger.log(`[${context}] ⏳ Nouvelle tentative dans ${delay}ms`);
                    Utilities.sleep(delay);
                    delay *= 2; // Backoff exponentiel
                }
            }
        }

        // Toutes les tentatives ont échoué
        this.logError(
            context,
            new Error(`Échec après ${maxRetries} tentatives: ${lastError.message}`),
            'CRITICAL',
            { originalError: lastError.message }
        );

        throw lastError;
    },

    /**
     * Valide une propriété de script et lève une erreur si manquante.
     * @param {string} propertyName - Nom de la propriété
     * @param {string} context - Contexte d'utilisation
     * @returns {string} Valeur de la propriété
     */
    requireProperty(propertyName, context = '') {
        const value = PropertiesService.getScriptProperties().getProperty(propertyName);

        if (!value) {
            const errorMsg = `Propriété de script manquante: ${propertyName}`;
            const fullContext = context ? `${context} - Configuration` : 'Configuration';

            this.logError(
                fullContext,
                new Error(errorMsg),
                'CRITICAL',
                {
                    propertyName,
                    helpMessage: `Veuillez configurer ${propertyName} dans Projet > Paramètres > Propriétés du script`
                }
            );

            throw new Error(`${errorMsg}. Veuillez configurer cette propriété dans les paramètres du script.`);
        }

        return value;
    },

    /**
     * Log une information de débogage (niveau INFO).
     * @param {string} context - Contexte
     * @param {string} message - Message
     * @param {object} metadata - Métadonnées optionnelles
     */
    logInfo(context, message, metadata = {}) {
        this.logError(context, message, 'INFO', metadata);
    },

    /**
     * Log un avertissement (niveau WARNING).
     * @param {string} context - Contexte
     * @param {string} message - Message
     * @param {object} metadata - Métadonnées optionnelles
     */
    logWarning(context, message, metadata = {}) {
        this.logError(context, message, 'WARNING', metadata);
    }
};

// Export pour compatibilité globale
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ErrorHandler;
}
