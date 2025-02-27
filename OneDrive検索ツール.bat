@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

:: 繧ｿ繧､繝医Ν險ｭ螳・
title OneDrive讀懃ｴ｢繝・・繝ｫ

:: 逕ｻ髱｢繧ｯ繝ｪ繧｢
cls
echo OneDrive讀懃ｴ｢繝・・繝ｫ
echo =====================================
echo.

:: 讀懃ｴ｢譁・ｭ怜・縺ｮ蜈･蜉帛女莉・
set /p SEARCH_TERM="讀懃ｴ｢縺励◆縺・枚蟄怜・繧貞・蜉帙＠縺ｦ縺上□縺輔＞: "

:: 讀懃ｴ｢蜈医ヵ繧ｩ繝ｫ繝縺ｮ險ｭ螳・
set ONEDRIVE_PATH=%USERPROFILE%\OneDrive

if not exist "%ONEDRIVE_PATH%" (
    echo OneDrive繝輔か繝ｫ繝縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ: %ONEDRIVE_PATH%
    goto :END
)

echo.
echo "%SEARCH_TERM%"繧呈､懃ｴ｢荳ｭ...縺雁ｾ・■縺上□縺輔＞...
echo.

:: 讀懃ｴ｢螳溯｡・
findstr /s /i /n /p "%SEARCH_TERM%" "%ONEDRIVE_PATH%\*.*" > "%TEMP%\onedrive_search_results.txt" 2>nul

:: 邨先棡陦ｨ遉ｺ
echo 讀懃ｴ｢邨先棡:
echo =====================================
type "%TEMP%\onedrive_search_results.txt"
echo =====================================
echo.
echo 讀懃ｴ｢縺悟ｮ御ｺ・＠縺ｾ縺励◆縲・
echo 邨先棡縺ｯ %TEMP%\onedrive_search_results.txt 縺ｫ繧ゆｿ晏ｭ倥＆繧後※縺・∪縺吶・

:END
pause
endlocal
