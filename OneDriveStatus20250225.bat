@echo off
REM �Ǘ��Ҍ����`�F�b�N�Ə��i
NET FILE 1>NUL 2>NUL
if not '%errorlevel%' == '0' (
    powershell -Command "Start-Process '%~f0' -Verb RunAs -Wait"
    exit /b
)

REM PowerShell�X�N���v�g�̃p�X��ݒ�
set "PS_SCRIPT=%~dp0OneDriveStatusCheckLatestVersion20250225.ps1"

REM PowerShell�X�N���v�g�𒼐ڎ��s�i�V���v���Ȏ��s���@�ɕύX�j
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

REM �X�N���v�g�������b�Z�[�W
echo.
echo �� �X�N���v�g�̎��s���������܂����BEnter�L�[�������ďI�����Ă��������B
pause > nul
