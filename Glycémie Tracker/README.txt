📊 Suivi Glycémie & Insuline — Excel Dashboard
🧭 Objectif du projet

Ce projet propose un fichier Excel complet permettant de :

Suivre les valeurs de glycémie dans le temps
Enregistrer les injections d’insuline
Calculer automatiquement les besoins en insuline (repas + correction)
Visualiser les données via des graphiques dynamiques

👉 Ce projet met en avant des compétences en modélisation Excel, automatisation, analyse de données et UX.

🏗️ Structure du fichier

Le classeur est organisé en 3 feuilles principales :

🟦 1. Saisie des données

Tableau principal utilisé pour la collecte des données.

Colonnes :

Date et heure : timestamp de la mesure
Glycémie : valeur mesurée
Quantité d’insuline : unités injectées
Raison de la piqûre : repas / correction / autre
Repas (glucides) : quantité de glucides consommés (en grammes)

✅ Bonnes pratiques :

Saisir les données en continu (sans lignes vides)
Respecter le format des colonnes
Utiliser des valeurs cohérentes (unités constantes)


🟨 2. Calculs & paramètres

Feuille dédiée aux calculs automatiques et à la configuration.

🔧 Paramètres fixes
Ratio insuline / glucides
Ratio de correction (baisse glycémique)
Cibles glycémiques (min / max)

✏️ Données utilisateur
Glycémie de départ
Glycémie cible
Quantité de glucides

⚙️ Calculs automatisés
Ratio de correction calculé
Moyenne, minimum et maximum glycémique
Insuline recommandée :
Correction
Repas
Total injection


📈 3. Visualisation

Feuille contenant un graphique dynamique permettant de :

Suivre l’évolution de la glycémie
Identifier les tendances
Détecter les épisodes d’hyper/hypoglycémie


⚙️ Fonctionnalités clés
📊 Dashboard Excel simple et lisible
🔢 Calcul automatique des doses d’insuline
📉 Statistiques (moyenne, min, max)
🎯 Gestion des objectifs glycémiques
📈 Graphique dynamique basé sur les données saisies
🛡️ Sécurisation des formules (cellules protégées)
🧠 Compétences démontrées

Ce projet illustre :

Excel avancé
Formules (SI, SIERREUR, ratios)
Tableaux structurés
Mise en forme conditionnelle
Data Analysis
Agrégation de données
Indicateurs statistiques
Suivi temporel
Data Visualization
Création de graphiques dynamiques
Lecture de tendances
Séparation saisie / calcul
Structure claire et maintenable

🚀 Améliorations possibles
Ajout de listes déroulantes (validation de données)
Calcul automatique avancé des doses personnalisées
Indicateurs médicaux (temps dans la cible, hypo/hyper)
Automatisation via macros (VBA)
Version mobile optimisée
Export PDF ou reporting automatique