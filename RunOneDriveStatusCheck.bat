@echo off
cls
echo ===================================================
echo  OneDrive �X�e�[�^�X�`�F�b�N�c�[��
echo ===================================================
echo.
echo ���̃c�[���́AOneDrive�̎g�p�󋵂��擾���A
echo ���|�[�g�𐶐����܂��B
echo.
echo ���s���̓E�B���h�E����Ȃ��ł�������...
echo.
echo ��Microsoft�F�؉�ʂ��\�����ꂽ��A
echo  �����̃A�J�E���g�Ń��O�C�����Ă��������B
echo ===================================================
echo.

cd /d "%~dp0"

echo PowerShell�X�N���v�g�����s���Ă��܂�...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0OneDriveStatusCheck.ps1"

echo.
if %errorlevel% neq 0 (
    echo �G���[���������܂����B�G���[���O���m�F���Ă��������B
    pause
    exit /b %errorlevel%
)

echo �������܂����I
echo.
echo �E�B���h�E����Ă��������B
pause
