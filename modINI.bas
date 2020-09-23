Attribute VB_Name = "modINI"
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                "GetPrivateProfileStringA" (ByVal lpApplicationName _
                As String, ByVal lpKeyName As Any, ByVal lpDefault _
                As String, ByVal lpReturnedString As String, ByVal _
                nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                "WritePrivateProfileStringA" (ByVal lpApplicationName _
                As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                ByVal lpFileName As String) As Long

Public Function sGetINI(sINIFile As String, sSection As String, sKey _
            As String, sDefault As String) As String
            
Dim sTemp As String * 256
Dim nLength As Integer

sTemp = Space$(256)
nLength = GetPrivateProfileString(sSection, sKey, sDefault, _
            sTemp, 255, sINIFile)
sGetINI = Left$(sTemp, nLength)
End Function

Public Sub WriteINI(sINIFile As String, sSection As String, _
            sKey As String, sValue As String)
            
Dim n As Integer
Dim sTemp As String

sTemp = sValue

'change  CR/LF by spaces
For n = 1 To Len(dValue)
    If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf _
    Then Mid$(sValue, n) = " "
Next n

n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
            
End Sub



