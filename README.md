# Bond Portfolio Analysis & Valuation Tool

Outil professionnel de gestion et d'analyse de portefeuille obligataire développé sous Excel et VBA.

## 📈 Fonctionnalités du Projet

### 1. Analyse Individuelle des Obligations
- Interface utilisateur dédiée à la saisie des caractéristiques obligataires (Nominal, coupon, maturité, fréquence, taux de marché).
- Fonctions VBA sur mesure :
  - `FluxObligation()` : Génération du tableau des flux de trésorerie (coupons et remboursement du nominal)[cite: 7].
  - `PrixObligation()` : Calcul de la valeur actuelle (prix plein / dirty price).
  - `TauxRendement()` : Calcul du taux de rendement actuariel (YTM) par dichotomie.
  - `Duration()` & `Convexite()` : Mesures de sensibilité du prix aux variations de taux.

### 2. Gestion et Agrégation du Portefeuille
- Univers de 15 obligations intégrant des données réelles de marché.
- Agrégation globale des flux financiers et calcul des KPI clés du portefeuille (Prix total, duration modifiée moyenne pondérée, convexité et rendement global).
- Projection barycentrique par piliers de maturité (de 1M à 30Y).
- Tableaux de bord graphiques : sensibilité par pilier de maturité et impact P&L suite à une variation de $\pm 1\%$ des taux.
- Automatisation complète via des boutons de mise à jour et génération dynamique de rapports par obligation.

---

## 🛠️ Utilisation et Installation

1. **Ouvrir le classeur Excel :**
   Assurez-vous d'activer les macros VBA (onglets développeur / sécurité des macros) à l'ouverture du fichier.
2. **Navigation dans le tableau de bord :**
   - Utilisez l'onglet `Interface_Obligation` pour analyser une obligation unitaire.
   - Consultez l'onglet `Dashboard_Obligation` et `Piliers` pour la vue d'ensemble du portefeuille et les graphiques de sensibilité.
3. **Actualisation :**
   Cliquez sur le bouton **"Mettre à jour"** intégré dans les feuilles pour recalculer l'ensemble des flux, prix et indicateurs de risque à la volée.

---

## 📂 Structure du Classeur (Onglets & Modules VBA)
- `Interface_Obligation` : Saisie et tarification unitaire.
- `Univers` & `Dashboard_Obligation` : Suivi global des 15 obligations et graphiques de valorisation.
- `Portefeuille_Flux` & `Piliers` : Centralisation des flux actualisés et analyse par piliers de maturité.
- Modules VBA associés : `Module Calculs`, `Module_Dashboard`, `Module_Graphes`, `Module Portefeuille`, `Module_Rapport`.
