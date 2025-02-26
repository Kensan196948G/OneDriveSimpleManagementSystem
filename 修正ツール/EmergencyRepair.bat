@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

echo =====================================================
echo   OneDrive�c�[���ً}�C�����[�e�B���e�B
echo =====================================================
echo.
echo ���̃c�[����OneDrive�^�p�c�[���ً̋}�C�����s���܂��B
echo ������������s�G���[���������Ă���ꍇ�Ɏg�p���Ă��������B
echo.

REM �Ǘ��Ҍ����`�F�b�N
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo �x��: ���̃X�N���v�g�͊Ǘ��Ҍ����Ȃ��Ŏ��s����Ă��܂��B
    echo �ꕔ�̑��삪���s����\��������܂��B
    echo �Ǘ��҂Ƃ��čĎ��s���邱�Ƃ𐄏����܂��B
    echo.
    pause
)

REM ���݂̃f�B���N�g���ƃv���W�F�N�g���[�g��ݒ�
set TOOL_DIR=%~dp0
cd /d %TOOL_DIR%
cd ..
set PROJECT_ROOT=%CD%

echo ��ƃf�B���N�g��: %PROJECT_ROOT%
echo.

echo �X�e�b�v 1: PowerShell�̎��s�|���V�[���m�F�E�ݒ蒆...
powershell -Command "Get-ExecutionPolicy" > %TEMP%\ExecPolicy.txt
set /p EXEC_POLICY=<%TEMP%\ExecPolicy.txt
del %TEMP%\ExecPolicy.txt
echo ���݂̎��s�|���V�[: %EXEC_POLICY%

if "%EXEC_POLICY%"=="Restricted" (
    echo ���s�|���V�[���ꎞ�I�ɕύX���܂�...
    powershell -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force"
    echo ���s�|���V�[��ύX���܂����B
)

echo.
echo �X�e�b�v 2: UTF-8�G���R�[�h�����m�F���Ă��܂�...
powershell -Command "[Console]::OutputEncoding.CodePage" > %TEMP%\Encoding.txt
set /p ENCODING=<%TEMP%\Encoding.txt
del %TEMP%\Encoding.txt

echo ���݂̃G���R�[�f�B���O: %ENCODING%
if not "%ENCODING%"=="65001" (
    echo �x��: �V�X�e���̃G���R�[�f�B���O��UTF-8�ł͂���܂���B
)

echo.
echo �X�e�b�v 3: �K�v�ȃt�H���_�̑��݊m�F�E�쐬...
if not exist "%PROJECT_ROOT%\logs" (
    mkdir "%PROJECT_ROOT%\logs"
    echo logs �t�H���_���쐬���܂���
)

if not exist "%PROJECT_ROOT%\�f�f" (
    mkdir "%PROJECT_ROOT%\�f�f"
    echo �f�f �t�H���_���쐬���܂���
)

if not exist "%PROJECT_ROOT%\�C���c�[��" (
    mkdir "%PROJECT_ROOT%\�C���c�[��"
    echo �C���c�[�� �t�H���_���쐬���܂���
)

echo.
echo �X�e�b�v 4: �X�N���v�g�G���R�[�f�B���O�̏C��...

echo [A] �K�{�X�N���v�g�̃G���R�[�f�B���O���C�����Ă��܂�...
set IMPORTANT_SCRIPTS=OneDriveStatusCheck.ps1

for %%s in (%IMPORTANT_SCRIPTS%) do (
    if exist "%PROJECT_ROOT%\%%s" (
        echo     %%s ���C����...
        powershell -NoProfile -Command "$content = [System.IO.File]::ReadAllText('%PROJECT_ROOT%\%%s', [System.Text.Encoding]::Default); [System.IO.File]::WriteAllText('%PROJECT_ROOT%\%%s', $content, [System.Text.Encoding]::UTF8)"
        echo     %%s ���C�����܂���
    ) else (
        echo     �x��: %%s ��������܂���
    )
)

echo [B] �o�b�`�t�@�C���̃G���R�[�f�B���O���C�����Ă��܂�...
for %%f in ("%PROJECT_ROOT%\*.bat") do (
    echo     %%~nxf ���C����...
    powershell -NoProfile -Command "$content = [System.IO.File]::ReadAllText('%%f', [System.Text.Encoding]::Default); [System.IO.File]::WriteAllText('%%f', $content, [System.Text.Encoding]::UTF8)"
)

echo.
echo �X�e�b�v 5: �G���R�[�f�B���O�C���c�[���̏���...
echo CreateFixerScript.ps1�����s���ďC���c�[���𐶐����Ă��܂�...
powershell -NoProfile -ExecutionPolicy Bypass -File "%TOOL_DIR%CreateFixerScript.ps1"

echo.
echo �X�e�b�v 6: PowerShell���W���[���̏�Ԃ��m�F...
powershell -NoProfile -Command "Get-Module -ListAvailable Microsoft.Graph.* | Format-Table Name, Version"

echo.
echo =====================================================
echo   �ً}�C���������������܂���
echo =====================================================
echo.
echo ��肪�������Ȃ��ꍇ�͈ȉ��������Ă�������:
echo.
echo 1. EncodingFixer.bat�����s���ăX�N���v�g���C��
echo 2. PowerShell���Ǘ��҂Ƃ��Ď��s���A�ȉ��̃R�}���h�����s:
echo    Install-Module Microsoft.Graph.Users -Force
echo    Install-Module Microsoft.Graph.Files -Force
echo    Install-Module Microsoft.Graph.Authentication -Force
echo.
echo 3. OneDriveStatusCheck.ps1�̎��s���Ɉȉ��̃p�����[�^��ǉ�:
echo    -ClearAuth
echo.

pause