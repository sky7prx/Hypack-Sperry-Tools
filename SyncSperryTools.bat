@echo off

setlocal

:: --- CONFIGURATION ---
:: 1. Enter the full path to the FOLDER containing git.exe inside GitPortable
set "GIT_BIN_PATH=%USERPROFILE%\Desktop\Sperry\PortableGit\bin"

:: 2. Enter the full path to your local Repository
set "REPO_PATH=%USERPROFILE%\Desktop\Sperry\Hypack-Sperry-Tools"
:: ---------------------

:: Add Git to the temporary PATH for this session only
set "PATH=%GIT_BIN_PATH%;%PATH%"

echo Navigating to: %REPO_PATH%
cd /d "%REPO_PATH%"

echo.
echo Pulling latest changes...
:: Using 'call' ensures the script continues if git-cmd initiates other processes
call git fetch --all
call git reset --hard origin/main

echo.
echo Update complete.
pause

set /p "copyfile=Copy the new file to the Desktop\Sperry folder? (y/n) "

if /I not "%copyfile%"=="y" goto theend

setlocal enabledelayedexpansion

set "dest=%USERPROFILE%\Desktop\Sperry"

echo Copy file to: %dest%

for %%F in ("%REPO_PATH%\Hypack to Sperry Route Converter.*") do (
    set "filename=%%~nxF"
    set "basename=%%~nF"
    set "ext=%%~xF"
    
    if not exist "%dest%\%%~nxF" (
        copy "%%F" "%dest%\"
    ) else (
        set /a count=1
        :loop
        set "newname=!basename!!count!!ext!"
        if exist "%dest%\!newname!" (
            set /a count+=1
            goto loop
        )
        copy "%%F" "%dest%\!newname!"
	echo Copied file to Desktop as !newname!
    )
)
pause

:theend
endlocal