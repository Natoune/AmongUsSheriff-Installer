@echo off

setlocal

@echo ---------------------------------------------------------------
@echo Telechargement de Among.Us.Sheriff.Mod.v1.24.v2021.6.30s.zip...
@echo ---------------------------------------------------------------
c:\windows\system32\cscript.exe extract.vbs
@echo ----------------------------
@echo Copie des fichiers du jeu...
@echo ----------------------------
xcopy "C:\Program Files\Epic Games\AmongUs" "%cd%\Game" /e /i

cd /d %~dp0

@echo ----------------------------
@echo Extraction de extract.zip...
@echo ----------------------------
Call :UnZipFile "%cd%\Game" "%cd%\extract.zip"
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%cd%\Among US Sheriff.lnk');$s.TargetPath='%cd%\Game\Among Us.exe';$s.Save()"
exit /b

:download
(echo src = "%~1"
echo Set v1 = CreateObject ("MSXML2.XMLHTTP"^)
echo Set v2  = CreateObject ("ADODB.Stream"^)
echo v1.open "GET", src, false
echo v1.send (^)
echo v2.open
echo v2.Type = 1
echo v2.Write v1.ResponseBody
echo v2.SaveToFile "%~2" ) >"%~dpn0.vbs"
cscript "%~dpn0.vbs"
del "%~dpn0.vbs" >nul
goto:eof

:UnZipFile <ExtractTo> <newzipfile>
set vbs="%temp%\_.vbs"
if exist %vbs% del /f /q %vbs%
>%vbs%  echo Set fso = CreateObject("Scripting.FileSystemObject")
>>%vbs% echo If NOT fso.FolderExists(%1) Then
>>%vbs% echo fso.CreateFolder(%1)
>>%vbs% echo End If
>>%vbs% echo set objShell = CreateObject("Shell.Application")
>>%vbs% echo set FilesInZip=objShell.NameSpace(%2).items
>>%vbs% echo objShell.NameSpace(%1).CopyHere(FilesInZip)
>>%vbs% echo Set fso = Nothing
>>%vbs% echo Set objShell = Nothing
cscript //nologo %vbs%
if exist %vbs% del /f /q %vbs%