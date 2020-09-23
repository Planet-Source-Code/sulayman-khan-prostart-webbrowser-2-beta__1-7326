Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'+=+=+=+=+=+=+=   INI   INI   INI
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private INIFileName As String
Private ret As String
Private RetLen As Long
'+=+=+=+=+=+=+=   INI INI INI INI

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public numberOfWindows As Integer ' Counts open windows

'+=+=+=+=+=+=+=   INI   INI   INI
Public Function GetValue(Section As String, Key As String, INIFileName As String) As Variant
ret = Space$(255)
RetLen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), INIFileName)
ret = left$(ret, RetLen)
GetValue = ret
End Function

Public Sub PutValue(Section As String, Key As String, Text As String, INIFileName As String)
WritePrivateProfileString Section, Key, Text, INIFileName
End Sub

Sub delSection(Section As String, INIFileName As String)
  'Deletes an *entire* [Section] and all its Entries
WritePrivateProfileString Section, 0&, 0&, INIFileName
End Sub

Function delKey(Section As String, Key As String, INIFileName As String)
    'deletes a key
    WritePrivateProfileString Section, Key, vbNullString, INIFileName
End Function
'+=+=+=+=+=+=+=   INI   INI   INI
