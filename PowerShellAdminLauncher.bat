@echo off
setlocal enabledelayedexpansion
chcp 932 > nul

REM PowerShell繧ｹ繧ｯ繝ｪ繝励ヨ繧堤ｮ｡逅・・ｨｩ髯舌〒螳溯｡後☆繧九◆繧√・繝ｩ繝ｳ繝√Ε繝ｼ
title OneDrive驕狗畑繝・・繝ｫ 邂｡逅・・ｮ溯｡後Λ繝ｳ繝√Ε繝ｼ
color 17

REM 邂｡逅・・ｨｩ髯舌メ繧ｧ繝・け
NET SESSION >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ==========================================
    echo   縺薙・繝・・繝ｫ縺ｯ邂｡逅・・ｨｩ髯舌〒螳溯｡後☆繧句ｿ・ｦ√′縺ゅｊ縺ｾ縺吶・
    echo   蜿ｳ繧ｯ繝ｪ繝・け縺九ｉ縲檎ｮ｡逅・・→縺励※螳溯｡後阪ｒ驕ｸ謚槭＠縺ｦ縺上□縺輔＞縲・
    echo ==========================================
    echo.
    echo 邨ゆｺ・☆繧九↓縺ｯ菴輔°繧ｭ繝ｼ繧呈款縺励※縺上□縺輔＞...
    pause >nul
    exit /b 1
)

:menu
cls
echo ==========================================
echo   OneDrive驕狗畑繝・・繝ｫ 邂｡逅・・ｮ溯｡後Λ繝ｳ繝√Ε繝ｼ
echo ==========================================
echo.
echo 邂｡逅・・ｨｩ髯舌〒螳溯｡後☆繧九せ繧ｯ繝ｪ繝励ヨ繧帝∈謚槭＠縺ｦ縺上□縺輔＞:
echo.
echo [1] OneDrive繧ｹ繝・・繧ｿ繧ｹ繝ｬ繝昴・繝井ｽ懈・
echo [2] OneDrive謗･邯夊ｨｺ譁ｭ繝・・繝ｫ螳溯｡・
echo [3] 譁・ｭ怜喧縺台ｿｮ豁｣繝・・繝ｫ襍ｷ蜍・
echo [0] 邨ゆｺ・
echo.
echo ==========================================
echo.

set /p OPTION=驕ｸ謚槭＠縺ｦ縺上□縺輔＞ (0-3): 

if "%OPTION%"=="1" goto :run_report
if "%OPTION%"=="2" goto :run_debug 
if "%OPTION%"=="3" goto :run_encoding_fix
if "%OPTION%"=="0" goto :end

echo 辟｡蜉ｹ縺ｪ驕ｸ謚槭〒縺吶・
timeout /t 2 >nul
goto :menu

:run_report
cls
echo 邂｡逅・・ｨｩ髯舌〒 OneDrive繧ｹ繝・・繧ｿ繧ｹ繝ｬ繝昴・繝井ｽ懈・繧貞ｮ溯｡後＠縺ｾ縺・..
echo.

set SCRIPT_PATH=%~dp0OneDriveStatusCheck.ps1

REM 譏守､ｺ逧・↓UTF-8繧ｨ繝ｳ繧ｳ繝ｼ繝・ぅ繝ｳ繧ｰ繧定ｨｭ螳壹＠縺ｦ繧ｹ繧ｯ繝ｪ繝励ヨ螳溯｡・
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8; ^
    $env:PYTHONIOENCODING = 'utf-8'; ^
    & '%SCRIPT_PATH%'"

echo.
echo 繝ｬ繝昴・繝井ｽ懈・縺悟ｮ御ｺ・＠縺ｾ縺励◆縲・
echo.
pause
goto :menu

:run_debug
cls
echo 邂｡逅・・ｨｩ髯舌〒 OneDrive謗･邯夊ｨｺ譁ｭ繝・・繝ｫ繧貞ｮ溯｡後＠縺ｾ縺・..
echo.

set DEBUG_SCRIPT_PATH=%~dp0OneDriveDebugLog.ps1

REM 譏守､ｺ逧・↓UTF-8繧ｨ繝ｳ繧ｳ繝ｼ繝・ぅ繝ｳ繧ｰ繧定ｨｭ螳壹＠縺ｦ繧ｹ繧ｯ繝ｪ繝励ヨ螳溯｡・
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8; ^
    $env:PYTHONIOENCODING = 'utf-8'; ^
    & '%DEBUG_SCRIPT_PATH%'"

echo.
echo 險ｺ譁ｭ縺悟ｮ御ｺ・＠縺ｾ縺励◆縲・
echo.
pause
goto :menu

:run_encoding_fix
cls
echo 邂｡逅・・ｨｩ髯舌〒 譁・ｭ怜喧縺台ｿｮ豁｣繝・・繝ｫ繧定ｵｷ蜍輔＠縺ｾ縺・..
echo.

REM 譁・ｭ怜喧縺台ｿｮ豁｣繝・・繝ｫ繧よ・遉ｺ逧・↓繧ｨ繝ｳ繧ｳ繝ｼ繝・ぅ繝ｳ繧ｰ繧呈欠螳壹＠縺ｦ襍ｷ蜍・
set GUI_TOOL=%~dp0CharacterEncodingFixer.ps1

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; ^
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8; ^
    $env:PYTHONIOENCODING = 'utf-8'; ^
    & '%GUI_TOOL%'"

echo.
echo 譁・ｭ怜喧縺台ｿｮ豁｣繝・・繝ｫ繧堤ｵゆｺ・＠縺ｾ縺励◆縲・
echo.
pause
goto :menu

:end
echo.
echo 繝・・繝ｫ繧堤ｵゆｺ・＠縺ｾ縺・..
timeout /t 2 >nul
endlocal
exit /b 0
