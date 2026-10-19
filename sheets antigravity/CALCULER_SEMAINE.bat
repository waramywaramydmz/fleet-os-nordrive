@echo off
chcp 65001 > nul
cls
echo ========================================================================
echo   🚗 FLEET OS — CALCUL HEBDOMADAIRE AUTOMATIQUE
echo   Conciliation Uber + Bolt - Export Virement Qonto
echo ========================================================================
echo.
echo Mode d'emploi :
echo 1. Assurez-vous d'avoir place vos 2 fichiers CSV Uber et Bolt dans ce dossier.
echo    (Par exemple : UBER_S39_REAL.csv et BOLT_S39_REAL.csv ou uber.csv et bolt.csv)
echo 2. Lancement du calcul automatique...
echo.

node calculer_semaine.cjs

echo.
echo ========================================================================
echo Fichiers mis a jour sur votre machine :
echo  - FLEET_OS_Nordrive_Reconciliation_Qonto.xlsx (Dashboard Excel complet)
echo  - EXPORT_QONTO_SEPA.csv (Fichier prete a importer dans Qonto)
echo  - DASHBOARD_SEMAINE.csv (Synthese CSV de la semaine)
echo ========================================================================
echo.
pause
