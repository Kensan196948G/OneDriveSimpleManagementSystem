@echo off
setlocal

:: �^�C�g���\��
cls
echo =====================================================
echo   PowerShell�X�N���v�g���������C���c�[��
echo =====================================================
echo.
echo ���������f�f�E�C��GUI���N�����Ă��܂�...
echo.

:: GUI�c�[���̃p�X
set GUI_TOOL=%~dp0CharacterEncodingFixer.ps1

:: �X�N���v�g���݃`�F�b�N
if not exist "%GUI_TOOL%" (
    echo �G���[: ���������C���c�[����������܂���:
    echo %GUI_TOOL%
    echo.
    echo ���s����ɂ͉����L�[�������Ă�������...
    pause >nul
    exit /b 1
)

:: PowerShell�̎��s�|���V�[���m�F
powershell -Command "Get-ExecutionPolicy" > "%TEMP%\pspolicy.txt"
set /p PS_POLICY=<"%TEMP%\pspolicy.txt"
del "%TEMP%\pspolicy.txt"

:: �K�v�ɉ����Ď��s�|���V�[���ꎞ�I�ɕύX
if /i "%PS_POLICY%"=="Restricted" (
    echo ���s�|���V�[���ꎞ�I�ɕύX���܂�...
    powershell -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force"
)

:: GUI�c�[�����s
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%GUI_TOOL%'"

endlocal
