Attribute VB_Name = "LIB_SysTray"
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Type NOTIFYICONDATA
    cbSize As Long
    mhWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Dim TheForm As NOTIFYICONDATA

Public Function SysTray(ByRef Pic1 As PictureBox)
  TheForm.cbSize = Len(TheForm)
  TheForm.mhWnd = Pic1.hwnd
  TheForm.hIcon = Pic1.Picture
  TheForm.uId = 1&
  TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  TheForm.ucallbackMessage = WM_MOUSEMOVE
  TheForm.szTip = App.Title & vbCrLf & "Version " & App.Major & "." & App.Minor & "." & App.Revision
  Shell_NotifyIcon NIM_ADD, TheForm
End Function

Function ModifyIcon(ByRef Pic1 As PictureBox)
  TheForm.cbSize = Len(TheForm)
  TheForm.mhWnd = Pic1.hwnd
  TheForm.hIcon = Pic1.Picture
  TheForm.uId = 1&
  TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  TheForm.ucallbackMessage = WM_MOUSEMOVE
  Shell_NotifyIcon NIM_MODIFY, TheForm
End Function

Public Sub CleanUpSystray()
  Shell_NotifyIcon NIM_DELETE, TheForm
End Sub

