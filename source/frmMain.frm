VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSysTray 
      Height          =   375
      Left            =   4260
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer tmrPoll 
      Interval        =   1000
      Left            =   4260
      Top             =   2640
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   1290
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblTopMost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   510
      TabIndex        =   3
      Top             =   2190
      Width           =   2775
   End
   Begin VB.Label lblStatInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DL: 0,0 kB/sec  UL: 0,0 kB/sec"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   690
      TabIndex        =   0
      Top             =   1470
      Width           =   3525
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuOpacity 
         Caption         =   "Opacity"
         Begin VB.Menu mnuPercent 
            Caption         =   "10%"
            Index           =   1
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "20%"
            Index           =   2
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "30%"
            Index           =   3
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "40%"
            Index           =   4
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "60%"
            Index           =   6
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "70%"
            Index           =   7
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "80%"
            Index           =   8
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "90%"
            Index           =   9
         End
         Begin VB.Menu mnuPercent 
            Caption         =   "100%"
            Index           =   10
         End
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuAutoStart 
         Caption         =   "Automatic startup"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_objIpHelper As CIpHelper
Private TransferRate                    As Single
Private TransferRate2                   As Single
Private Xstart As Long
Private Ystart As Long

Private Sub Form_DblClick()
'  If Me.Caption <> "" Then
'    Me.Caption = ""
'  Else
'    Me.Caption = GG_PROGRAM_NAME & " " & App.Major & "." & App.Minor & "." & App.Revision
'  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    mnuHide_Click
  ElseIf KeyCode = vbKeyF1 Then
    mnuAbout_Click
  End If
End Sub

Private Sub Form_Load()
  Dim iAutoStart As Integer
  If App.PrevInstance = True Then
    End
  End If
  DoEvents
  Set m_objIpHelper = New CIpHelper
  
  mnuMain.Visible = False
  gg_UserINIFileName = App.Path & "\" & GG_REGKEYVALUE & ".ini"
  SysTray picSysTray
  
  Me.Top = Val(ReadINI("formposition", "maintop", "0", gg_UserINIFileName))
  Me.Left = Val(ReadINI("formposition", "mainleft", "0", gg_UserINIFileName))
  Me.Width = Val(ReadINI("formposition", "mainWidth", "3000", gg_UserINIFileName))
  Me.Height = Val(ReadINI("formposition", "mainHeight", "1000", gg_UserINIFileName))
  iAutoStart = Val(ReadINI("info", "AutoStart", "0", gg_UserINIFileName))
  If Me.Left > Screen.Width Then Me.Left = 0
  If Me.Top > Screen.Height Then Me.Top = 0
  If iAutoStart = 0 Then
    If MsgBox("Do you want """ & App.Title & """ to execute each time Windows starts?", vbQuestion + vbYesNo) = vbYes Then
      mnuAutoStart_Click
    End If
  Else
    Dim sVal As String
    LeggiChiaveStringa GG_REGKEY, GG_REGKEYVALUE, sVal
    mnuAutoStart.Checked = (sVal <> "")
  End If
  'Me.Caption = Trim(ReadINI("formposition", "mainCaption", GG_PROGRAM_NAME & " " & App.Major & "." & App.Minor & "." & App.Revision, gg_UserINIFileName))
  mnuPercent_Click Val(ReadINI("formposition", "alpha", "10", gg_UserINIFileName))
  
  SetTopMostWindow Me.hwnd, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = vbLeftButton Then
    Xstart = X
    Ystart = y
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = vbLeftButton Then
    If Me.Left + (X - Xstart) >= 0 Then
      Me.Left = Me.Left + (X - Xstart)
    Else
      Me.Left = 0
    End If
    
    If Me.Top + (y - Ystart) >= 0 Then
      Me.Top = Me.Top + (y - Ystart)
    Else
      Me.Top = 0
    End If
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = 2 Then
    PopupMenu mnuMain
  End If
End Sub

Private Sub Form_Resize()
  If Me.ScaleWidth < 500 Then
    Me.Width = 600
  End If
  If Me.ScaleHeight < 500 Then
    Me.Height = 600
  End If
  
  '
  ' Set lblStatInfo.Height
  '
  LabelFontResize lblStatInfo, Me.ScaleWidth - (GG_MARGIN * 2)
  lblStatInfo.Left = GG_MARGIN
  lblStatInfo.Width = Me.ScaleWidth - (GG_MARGIN * 2)
  If lblStatInfo.Visible Then
    lblStatInfo.Top = Me.ScaleHeight - lblStatInfo.Height - GG_MARGIN
  Else
    lblStatInfo.Top = Me.ScaleHeight
  End If
  
  picGraph.Top = GG_MARGIN
  picGraph.Left = GG_MARGIN
  picGraph.Width = lblStatInfo.Width
  picGraph.Height = lblStatInfo.Top - picGraph.Top - GG_MARGIN
  
'  LabelFontResize lblInterface, picGraph.Width
  
  lblTopMost.Left = -GG_MARGIN
  lblTopMost.Top = -GG_MARGIN
  lblTopMost.Width = Me.Width + GG_MARGIN * 2
  lblTopMost.Height = Me.Height + GG_MARGIN * 2
  
  tmrPoll_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetTopMostWindow Me.hwnd, False
  CleanUpSystray
  Call WriteINI("formposition", "maintop", Me.Top, gg_UserINIFileName)
  Call WriteINI("formposition", "mainleft", Me.Left, gg_UserINIFileName)
  Call WriteINI("formposition", "mainWidth", Me.Width, gg_UserINIFileName)
  Call WriteINI("formposition", "mainHeight", Me.Height, gg_UserINIFileName)
  Call WriteINI("formposition", "mainCaption", Me.Caption & " ", gg_UserINIFileName)
  Call WriteINI("info", "AutoStart", IIf(mnuAutoStart.Checked, "1", "2"), gg_UserINIFileName)
  Dim i As Integer
  For i = 1 To 10
    If mnuPercent(i).Checked Then
      Call WriteINI("formposition", "alpha", CStr(i), gg_UserINIFileName)
      Exit For
    End If
  Next
End Sub

Private Sub lblTopMost_DblClick()
  Form_DblClick
End Sub

Private Sub lblTopMost_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  Form_MouseDown Button, Shift, X, y
End Sub

Private Sub lblTopMost_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  Form_MouseMove Button, Shift, X, y
End Sub

Private Sub lblTopMost_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = 2 Then
    PopupMenu mnuMain
  End If
End Sub

Private Sub lblInterface_DblClick()
  Form_DblClick
End Sub

Private Sub lblInterface_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  Form_MouseDown Button, Shift, X, y
End Sub

Private Sub lblInterface_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  Form_MouseMove Button, Shift, X, y
End Sub

Private Sub lblInterface_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = 2 Then
    PopupMenu mnuMain
  End If
End Sub

Private Sub mnuAbout_Click()
  frmAbout.AboutInfo , , , , Me
End Sub

Private Sub mnuAutoStart_Click()
  On Error Resume Next
  If mnuAutoStart.Checked Then
    mnuAutoStart.Checked = False
    'remove from autostart
    CancellaValore GG_REGKEY, GG_REGKEYVALUE
  Else
    'add to autostart
    ScriviChiaveStringa GG_REGKEY, GG_REGKEYVALUE, App.Path & "\" & App.EXEName & ".exe"
    mnuAutoStart.Checked = True
  End If
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuHide_Click()
  If mnuHide.Caption = "Hide" Then
    mnuHide.Caption = "Show"
    Me.Hide
  Else
    mnuHide.Caption = "Hide"
    Me.Show
  End If
End Sub

Private Sub mnuPercent_Click(Index As Integer)
  Static iIndexChecked As Integer
  
  If iIndexChecked <> 0 Then mnuPercent(iIndexChecked).Checked = False
  SetSemiTransparent Me.hwnd, CByte((255 * Index) / 10)
  mnuPercent(Index).Checked = True
  iIndexChecked = Index
End Sub

Private Sub picGraph_DblClick()
  Form_DblClick
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  Form_MouseDown Button, Shift, X, y
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  Form_MouseMove Button, Shift, X, y
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = 2 Then
    PopupMenu mnuMain
  End If
End Sub

Private Sub picSysTray_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  Static Rec As Boolean, Msg As Long
  
  Msg = X / Screen.TwipsPerPixelX
  If Rec = False Then
    Rec = True
    Select Case Msg
      Case WM_LBUTTONDBLCLK:
      Case WM_LBUTTONDOWN:
        mnuHide_Click
      Case WM_LBUTTONUP:
      Case WM_RBUTTONDBLCLK:
      Case WM_RBUTTONDOWN:
      Case WM_RBUTTONUP:
        PopupMenu mnuMain
    End Select
    Rec = False
  End If
End Sub

Private Sub tmrPoll_Timer()
  tmrPoll.Enabled = False
  On Error GoTo ErrH
  Dim objInterface        As CInterface
  Static lngBytesRecv     As Long
  Static lngBytesSent     As Long
  Dim lIn As Long, lOut As Long
  
  Set objInterface = m_objIpHelper.Interfaces(1)
  lIn = m_objIpHelper.BytesReceived - lngBytesRecv - 3296
  lOut = m_objIpHelper.BytesSent - lngBytesSent - 3296
  If lIn < 0 Then lIn = 0
  If lOut < 0 Then lOut = 0
  
  lblStatInfo.Caption = "DL: " & GetTransferRate(lIn) & "/sec  UL: " & GetTransferRate(lOut) & "/sec"
  
  picGraph.ScaleMode = 3      ' Set ScaleMode to pixels.
  'Lines picGraph.hDc, 1, 1, picGraph.ScaleWidth, picGraph.ScaleHeight, &H0
  DrawUsage picGraph, lIn, lOut
  lngBytesRecv = m_objIpHelper.BytesReceived
  lngBytesSent = m_objIpHelper.BytesSent
  DoEvents
  tmrPoll.Enabled = True
Exit Sub
ErrH:
  tmrPoll.Enabled = True
  Debug.Print Err.Description
End Sub

Function GetTransferRate(pDiff As Long) As String
  Dim d As Double
  
  ' bytes
'  If pDiff < 1000 Then
'    GetTransferRate = Trim(Format(pDiff, "##0")) & " b"
'    Exit Function
'  End If
  
  ' kbytes
  d = pDiff / 1024
  If d < 1024 Then
    GetTransferRate = Trim(Format(d, "#,##0.00")) & " Kb"
    Exit Function
  End If
  
  ' Mbytes
  d = pDiff / 1024
  GetTransferRate = Trim(Format(d, "#,##0.00")) & " Mb"
End Function

'Function GetInterfaceType(pIntTyp As InterfaceTypes) As String
'  Select Case pIntTyp
'    Case MIB_IF_TYPE_ETHERNET: GetInterfaceType = "Ethernet"
'    Case MIB_IF_TYPE_FDDI: GetInterfaceType = "FDDI"
'    Case MIB_IF_TYPE_LOOPBACK: GetInterfaceType = "Loopback"
'    Case MIB_IF_TYPE_OTHER: GetInterfaceType = "Other"
'    Case MIB_IF_TYPE_PPP: GetInterfaceType = "PPP"
'    Case MIB_IF_TYPE_SLIP: GetInterfaceType = "SLIP"
'    Case MIB_IF_TYPE_TOKENRING: GetInterfaceType = "TokenRing"
'  End Select
'End Function
'
'
