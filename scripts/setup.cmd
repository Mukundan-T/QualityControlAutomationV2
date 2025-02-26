@Echo off

set root_dir=..
set app_dir=%root_dir%\app
set folder=%~dp0..\

echo Installing requirements ...
python -m pip install --upgrade pip
pip install -r %root_dir%\requirements.txt

echo:
echo Press enter to select a location for Program Icon:
pause>Nul
python src\select_dir.py > %app_dir%\src\assets\homeDir.txt
set /p IconDir=<%app_dir%\src\assets\homeDir.txt

:: Default if userdoes not select a file in explorer
IF "%IconDir%"=="C:\" set IconDir=%userprofile%\Desktop

set Shortcut="%iconDir%\Quality Control.lnk"

:: Creates a vb executable to create .lnk file on the desktop
IF EXIST "%Shortcut%" (
    echo Desktop Shortcut Located ...
    echo:
) ELSE (
echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo sLinkFile = %Shortcut% >> CreateShortcut.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateShortcut.vbs
echo oLink.TargetPath = "%folder%\app\src\main.pyw%" >> CreateShortcut.vbs
echo oLink.Description = "QualityControl v1.0" >> CreateShortcut.vbs
echo oLink.IconLocation = "%folder%\app\src\assets\icon.ico" >> CreateShortcut.vbs
echo oLink.Save >> CreateShortcut.vbs
cscript CreateShortcut.vbs>Nul
del CreateShortcut.vbs
)


echo You're all set up! Please Proceed.
pause
