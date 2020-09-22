Attribute VB_Name = "LIB_INI"
Option Explicit
'INI File Functions...
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  
Function ReadINI(sSection As String, sKeyName As String, sDefaultValue As String, sINIFileName As String) As String
  On Local Error Resume Next
  Dim sRet As String
  
  sRet = String(255, Chr(0))
  ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, sDefaultValue, sRet, Len(sRet), sINIFileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
  On Local Error Resume Next
  Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
  WriteINI = (Err.Number = 0)
End Function

