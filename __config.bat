@ECHO OFF
title Folder lib

set "folderPath=lib"

:MENU
echo Do you want to (H)ide or (S)how the folder? (H/S):
set /p "cho=>"
if /I %cho%==H goto HIDE
if /I %cho%==S goto SHOW
echo Invalid choice. Please enter H or S.
goto MENU

:HIDE
attrib +h +s "%folderPath%"
echo Folder is now hidden and will not be visible even with hidden files shown.
goto MENU

:SHOW
echo Enter password to unlock folder:
set /p "pass=>"
if NOT %pass%==123.456+ goto FAIL
attrib -h -s "%folderPath%"
echo Folder Unlocked successfully.
goto MENU

:FAIL
echo Invalid password. Please try again.
goto MENU