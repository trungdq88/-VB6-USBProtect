Attribute VB_Name = "Module1"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'Ghi file
Function WriteIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sText As String) As Boolean
Dim i As Integer
On Error GoTo sWriteIniFileError
i = WritePrivateProfileString(sSection, sItem, sText, sIniFileName)
WriteIniFile = True
Exit Function
sWriteIniFileError:
WriteIniFile = False
End Function
'Doc file
Function ReadIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sDefault As String) As String
Dim iRetAmount As Integer
Dim sTemp As String
sTemp = String$(50, 0)
iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 50, sIniFileName)
sTemp = Left$(sTemp, iRetAmount)
ReadIniFile = sTemp
End Function

