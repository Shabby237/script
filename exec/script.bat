@echo off
call:download "http://www.google.fr/images/srpr/logo3w.png" "%~dp0logo3w.png"
exit /b

:download
(echo src = "%~1"
echo Set v1 = CreateObject("MSXML2.XMLHTTP"^)
echo Set v2  = CreateObject("ADODB.Stream"^)
echo v1.open "GET", src, false
echo v1.send(^)
echo v2.open
echo v2.Type = 1
echo v2.Write v1.ResponseBody
echo v2.SaveToFile "%~2") >"%~dpn0.vbs"

cscript "%~dpn0.vbs"

@REM del "%~dpn0.vbs" >nul

goto:eof