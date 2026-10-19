@echo off
chcp 65001 > nul
echo ============================================================
echo   FLEET OS — GENERATION DU SYSTEME GOOGLE SHEETS & QONTO
echo   Conciliation Uber + Bolt - Retrocession SEPA
echo ============================================================
echo.
echo Generation du classeur multi-onglets, du script Google Apps
echo et du fichier CSV conforme Qonto...
echo.

node calculer_semaine.cjs

echo.
echo ============================================================
echo   CLASSEUR & SCRIPTS GENERES AVEC SUCCES !
echo ============================================================
echo.
echo Fichiers prets :
echo  - FLEET_OS_Nordrive_Reconciliation_Qonto.xlsx (A importer dans Google Sheets)
echo  - GoogleAppsScript_FLEET_OS.js (Script 1-clic pour Extensions > Apps Script)
echo  - EXPORT_QONTO_SEPA.csv (Pret pour import direct sur Qonto)
echo  - FLEET_OS_Portail_Google_Sheets.html (Portail interactif)
echo.
echo Ouverture du portail interactif...
start FLEET_OS_Portail_Google_Sheets.html
pause
