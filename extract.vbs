Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetAbsolutePathName(".")

URL = "https://github.com/Woodi-dev/Among-Us-Sheriff-Mod/releases/download/v1.24_2021.6.30s/Among.Us.Sheriff.Mod.v1.24.v2021.6.30s.zip"
objWinHttp.open "GET", URL, False
objWinHttp.send ""

SaveBinaryData CurrentDirectory & "\extract.zip",objWinHttp.responseBody

Function SaveBinaryData(FileName, Data)

' adTypeText for binary = 1
Const adTypeText = 1
Const adSaveCreateOverWrite = 2

' Create Stream object
Dim BinaryStream
Set BinaryStream = CreateObject("ADODB.Stream")

' Specify stream type - we want To save Data/string data.
BinaryStream.Type = adTypeText

' Open the stream And write binary data To the object
BinaryStream.Open
BinaryStream.Write Data

' Save binary data To disk
BinaryStream.SaveToFile FileName, adSaveCreateOverWrite

End Function