@echo off
cls
echo ===================================================
echo  OneDrive 繧ｹ繝・・繧ｿ繧ｹ繝√ぉ繝・け繝・・繝ｫ
echo ===================================================
echo.
echo 縺薙・繝・・繝ｫ縺ｯ縲＾neDrive縺ｮ菴ｿ逕ｨ迥ｶ豕√ｒ蜿門ｾ励＠縲・
echo 繝ｬ繝昴・繝医ｒ逕滓・縺励∪縺吶・
echo.
echo 螳溯｡御ｸｭ縺ｯ繧ｦ繧｣繝ｳ繝峨え繧帝哩縺倥↑縺・〒縺上□縺輔＞...
echo.
echo 窶ｻMicrosoft隱崎ｨｼ逕ｻ髱｢縺瑚｡ｨ遉ｺ縺輔ｌ縺溘ｉ縲・
echo  閾ｪ蛻・・繧｢繧ｫ繧ｦ繝ｳ繝医〒繝ｭ繧ｰ繧､繝ｳ縺励※縺上□縺輔＞縲・
echo ===================================================
echo.

cd /d "%~dp0"

echo PowerShell繧ｹ繧ｯ繝ｪ繝励ヨ繧貞ｮ溯｡後＠縺ｦ縺・∪縺・..
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0OneDriveStatusCheck.ps1"

echo.
if %errorlevel% neq 0 (
    echo 繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆縲ゅお繝ｩ繝ｼ繝ｭ繧ｰ繧堤｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲・
    pause
    exit /b %errorlevel%
)

echo 螳御ｺ・＠縺ｾ縺励◆・・
echo.
echo 繧ｦ繧｣繝ｳ繝峨え繧帝哩縺倥※縺上□縺輔＞縲・
pause
