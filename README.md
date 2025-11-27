# KPI NANCOMCY - Dashboard Automatisé Marketing

> Système d'automatisation pour collecter et agréger les KPIs marketing depuis 9+ sources de données.

## 🏗️ Architecture

### Fichiers Principaux
- **`Utils.js`** - Fonctions utilitaires partagées (tokens, normalization, dates, Monday GraphQL)
- **`Config.js`** - Configuration centralisée (constantes, URLs, IDs)
- **`Tests.js`** - Tests unitaires
- **`Site Internet.js`** - KPIs du site web (GA4, GSC, Matomo)
- **`Google Ads.js`** - KPIs Google Ads  
- **`Meta Ads.js`** - KPIs Meta/Facebook/Instagram
- **`GMB.js`** - KPIs Google Business Profile
- **`Sheet_Structures.md`** - Documentation structure des feuilles

### Sources de Données
1. **Google Analytics 4** - Sessions (jusqu'à oct 2025), durée, rebond
2. **Matomo** ⭐ - Sessions (à partir de nov 2025), avec segmentation
3. **Google Search Console** - Impressions, clics organiques
4. **Google Ads API** - Performances publicitaires
5. **Meta Marketing API** - Données Facebook/Instagram
6. **Google Business Profile** - Vues, interactions GMB
7. **Magnetis** - Appels téléphoniques
8. **Paperform** - Soumissions de formulaires
9. **Monday.com** - Leads et formulaires

## ⚙️ Configuration

### Propriétés de Script Requises

1. Ouvrir **Apps Script Editor**
2. **Projet** > **Paramètres** > **Propriétés du script**
3. Ajouter les propriétés suivantes :

| Propriété | Description | Requis |
|-----------|-------------|--------|
| `MATOMO_TOKEN` | Token d'authentification Matomo | ✅ |
| `MATOMO_SEGMENT_ID` | ID du segment Matomo | ⚪ Optionnel |
| `MAGNETIS_API_KEY` | Clé API Magnetis | ✅ |
| `PAPERFORM_TOKEN` | Token API Paperform | ✅ |
| `MONDAY_TOKEN` | Token API Monday.com | ✅ |
| `META_TOKEN` | Token long-lived Meta Marketing API | ✅ |
| `CLIENT_ID` | OAuth Client ID (Google Ads) | ✅ |
| `CLIENT_SECRET` | OAuth Client Secret (Google Ads) | ✅ |

## 🚀 Utilisation

### Fonctions Principales

#### Full History (21 derniers mois)
```javascript
run_Site_FullHistory()      // Site Internet
run_GAds_FullHistory()      // Google Ads
run_GBM_FullHistory()       // GMB
```

#### Last Month (N-1)
```javascript
run_Site_AddLastMonth()     // Site Internet
Ads_run_All_AddLastMonth()  // Google Ads
run_GBM_AddLastMonth()      // GMB
```

#### Current Month
```javascript
run_Site_CurrentMonth()     // Site Internet ⭐
Ads_run_All_CurrentMonth()  // Google Ads
```

### Déclencheurs Recommandés

| Fonction | Fréquence | Horaire suggéré |
|----------|-----------|-----------------|
| `run_All_FullHistory()` | 1x/mois | 1er du mois, 2h |
| `Ads_run_All_AddLastMonth()` | 1x/mois | 5e jour, 6h |
| `Ads_run_All_CurrentMonth()` | Quotidien | 6h |

## 🧪 Tests

Exécuter les tests unitaires :

```javascript
runAllTests()
```

**Tests inclus** :
- ✅ Normalisation des en-têtes (`Utils_normHeader`)
- ✅ Conversion des dates (`Utils_sheetCellToYYYYMM`)
- ✅ Parsing Matomo (`SITE_parseMatomoVisitsResponse_`)
- ✅ Comparaison de texte (`Utils_textEqualsAny`, `Utils_textIncludesAny`)
- ✅ Conversion date française (`Utils_monthKeyToFr`)

## 🔧 Maintenance

### Ajouter une Nouvelle Métrique

1. Ajouter la colonne dans le Google Sheet
2. Mettre à jour `Sheet_Structures.md`
3. Modifier `_findCols_()` du fichier approprié
4. Ajouter logique dans `_writeMonthRow_()`
5. Tester avec `runAllTests()`

### Ajouter une Nouvelle Source de Données

1. Créer les fonctions de collecte
2. Ajouter le token dans les propriétés du script (via `Utils_getProperty`)
3. Intégrer dans `run_XXX_FullHistory()`
4. Documenter ici

## 🐛 Dépannage

### Erreur "Property 'XXX_TOKEN' not found"
**Solution** : Vérifier que la propriété est configurée dans Projet > Paramètres > Propriétés du script.

### Timeout lors de l'exécution
**Solutions** :
- Réduire la fenêtre de dates
- Exécuter les fonctions séparément par source
- Vérifier les logs pour identifier la source lente

### Données manquantes
**Checklist** :
1. Vérifier les logs (`View > Logs`)
2. S'assurer que les autorisations API sont accordées
3. Vérifier que les segments Matomo sont corrects
4. Confirmer que les tokens n'ont pas expiré

### Visiteurs du Current Month = 0 (Nov 2025+)
**Diagnostic** :
- Vérifier que `MATOMO_SEGMENT_ID` est défini
- Consulter les logs Matomo dans la sortie :
  ```
  [Matomo Visits] Période: 2025-11-01,2025-11-27
  [Matomo Parse] 2025-11: 456 visites
  ```
- Si "0 visites", vérifier le segment dans Matomo

## 📊 Migration Matomo (Important)

**Date de bascule** : 1er novembre 2025

À partir de cette date, les **visites** sont collectées depuis **Matomo** (avec segment) au lieu de **GA4**.

```javascript
const matomoCutoffDate = new Date(2025, 10, 1); // 1er Novembre 2025
```

**Métriques concernées** :
- ✅ **Sessions/Visites** : Matomo (avec segment)
- ⚪ **Durée moyenne** : GA4 (pour l'instant)
- ⚪ **Taux de rebond** : GA4 (pour l'instant)

## 📝 Changelog

### v2.0 (Nov 2025) - Stabilisation & Matomo
- ✅ **BUG FIX** : Correction parsing Matomo pour Current Month
- ✅ Ajout `Config.js` - Configuration centralisée
- ✅ Ajout `Tests.js` - Tests unitaires
- ✅ Ajout `Utils_getProperty()` - Gestion tokens avec fallback
- ✅ Ajout `Utils_mondayGraphQL()` - Monday centralisé
- ✅ Correction typo `PAPARFORM_TOKEN` → `PAPERFORM_TOKEN`
- ✅ Suppression ~60 lignes de code dupliqué

### v1.0 (Oct 2025) - Refactoring Initial
- Création `Utils.js` avec fonctions partagées
- Normalisation des en-têtes
- Protection des formules Excel

## 🤝 Contribution

1. Créer une branche pour modifications
2. Ajouter tests pour nouvelles fonctionnalités
3. Exécuter `runAllTests()` avant commit
4. Documenter changements dans README

## 📄 Licence

Usage interne NANCOMCY.
