@REM clasp-helper.bat - Gestion Événements SDIS 66
@echo off

REM Deployment ID fixe - ne change jamais, l'URL reste stable
set DEPLOY_ID=AKfycbzJNLnl5AsdE2I9vdVIPGzPSpxNqtzkTejn7hCig5qFan3-IB4f8eHpk78EYI5xFg3FgA

REM Usage: clasp-helper.bat [push|deploy|pull|open]

if "%1"=="push" (
    echo === Pushing to Google Apps Script ===
    cmd /c "clasp push --force"
    echo Done!
    goto :eof
)
if "%1"=="deploy" (
    echo === Deploying (same URL)... ===
    cmd /c "clasp deploy -i %DEPLOY_ID% -d stable"
    echo Done!
    goto :eof
)
if "%1"=="pull" (
    echo === Pulling from Google Apps Script ===
    cmd /c "clasp pull"
    echo Done!
    goto :eof
)
if "%1"=="pushdeploy" (
    echo === Push + Deploy (same URL) ===
    cmd /c "clasp push --force"
    cmd /c "clasp deploy -i %DEPLOY_ID% -d stable"
    echo Done!
    goto :eof
)
if "%1"=="open" (
    echo === Opening in browser ===
    cmd /c "clasp open"
    goto :eof
)

echo Gestion Evenements - Clasp Helper
echo ==================================
echo Usage: clasp-helper.bat [command]
echo   push        - Push vers Google Apps Script
echo   pull        - Pull depuis Google Apps Script
echo   deploy      - Deployer la webapp
echo   pushdeploy  - Push + Deploy
echo   open        - Ouvrir dans le navigateur
