@echo off
setlocal

echo ======================================================
echo  OneDrive繧ｹ繝・・繧ｿ繧ｹ繝ｬ繝昴・繝・(隱崎ｨｼ繝ｪ繧ｻ繝・ヨ迚・
echo ======================================================
echo.
echo 縺薙・繧ｹ繧ｯ繝ｪ繝励ヨ縺ｯ隱崎ｨｼ繧ｭ繝｣繝・す繝･繧偵け繝ｪ繧｢縺励※縺九ｉ
echo OneDrive繧ｹ繝・・繧ｿ繧ｹ繝ｬ繝昴・繝医ｒ逕滓・縺励∪縺吶・
echo.
echo 縲窟ccess denied縲阪∪縺溘・縲後い繧ｯ繧ｻ繧ｹ諡貞凄縲阪お繝ｩ繝ｼ縺・
echo 逋ｺ逕溘＠縺溷ｴ蜷医↓螳溯｡後＠縺ｦ縺上□縺輔＞縲・
echo.
echo ======================================================
echo.

set POWERSHELL=powershell.exe -NoProfile -ExecutionPolicy Bypass

echo 隱崎ｨｼ繧ｭ繝｣繝・す繝･繧偵け繝ｪ繧｢縺励※縺・∪縺・..
%POWERSHELL% -Command "Disconnect-MgGraph -ErrorAction SilentlyContinue"

echo.
echo 繧ｨ繝ｩ繝ｼ逋ｺ逕滓凾縺ｫ繧ゅΘ繝ｼ繧ｶ繝ｼ繧偵せ繧ｭ繝・・縺励※蜃ｦ逅・ｒ邯咏ｶ壹＠縺ｾ縺・
echo.

%POWERSHELL% -File "C:\kitting\OneDrive驕狗畑繝・・繝ｫ\OneDriveStatusCheck.ps1" -SkipErrorUsers $true -ClearAuth

echo.
echo 蜃ｦ逅・′螳御ｺ・＠縺ｾ縺励◆縲・
pause
