VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00000080&
      Height          =   2070
      Left            =   2715
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmAbout.frx":0000
      Top             =   1230
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Frame frameInfo 
      Height          =   1155
      Left            =   660
      TabIndex        =   2
      Top             =   1155
      Width           =   1950
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   1785
      End
      Begin VB.TextBox txtLicense 
         BackColor       =   &H00E0E0E0&
         Height          =   540
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   405
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2340
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.Image imgMouseOver 
      Height          =   480
      Left            =   255
      Picture         =   "frmAbout.frx":00B6
      Top             =   3315
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   165
      Picture         =   "frmAbout.frx":03C0
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblProgramAuthor 
      AutoSize        =   -1  'True
      Caption         =   "Myself"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1650
      TabIndex        =   6
      Top             =   615
      Width           =   690
   End
   Begin VB.Label lblProgramVersion 
      AutoSize        =   -1  'True
      Caption         =   "1.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2370
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblProgramName 
      AutoSize        =   -1  'True
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1650
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "by:"
      Height          =   195
      Left            =   1335
      TabIndex        =   3
      Top             =   645
      Width           =   210
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API declarations for local procedure: Exec
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'-------------------------------------------------------------------------------
' PROCEDURE   : Exec
' DESCRIPTION : Opens the specified document in the default associated program,
'   or executes the specified file
' RETURN VALUE: ???
' EXAMPLE     :
'   exec("C:\WINNT\Notepad.exe")
'   exec("C:\WINNT\win.ini")
'   exec("mailto:DGiabbai@hotpop.com")
' Copyright © 2002 Daniele Giabbai
'-------------------------------------------------------------------------------
Private Function Exec(pDocumentName As String) As Long
  Exec = ShellExecute(0, "open", pDocumentName, vbNullString, vbNullString, vbNull)
End Function

Public Sub AboutInfo(Optional p_programName As String, _
    Optional p_programVersion As String, _
    Optional p_programAuthor As String = "Daniele Giabbai", _
    Optional p_info As String, _
    Optional p_Form As Form, _
    Optional p_modal As Boolean = True, _
    Optional p_url As String = "http://www.geocities.com/giabbai", _
    Optional p_email As String = "mailto:DGiabbai@hotpop.com")
  
  lblProgramName.Caption = IIf(p_programName = "", Replace$(App.Title, "&", "&&"), p_programName)
  lblProgramName.ToolTipText = p_url
  lblProgramVersion.Caption = IIf(p_programVersion = "", App.Major & "." & App.Minor & "." & App.Revision, p_programVersion)
  lblProgramAuthor.Caption = p_programAuthor
  lblProgramAuthor.ToolTipText = p_email
  txtInfo.Text = IIf(p_info = "", App.Title, p_info)
  GNUGeneralPublicLicense txtLicense
  lblProgramVersion.Left = lblProgramName.Left + lblProgramName.Width + 30
  
  If Not p_Form Is Nothing Then
    Me.Icon = p_Form.Icon
    Image1.Picture = p_Form.Icon
    Me.Show IIf(p_modal, vbModal, vbModeless), p_Form
  Else
    Me.Icon = LoadPicture()
    Me.Show IIf(p_modal, vbModal, vbModeless)
  End If
End Sub

Sub GNUGeneralPublicLicense(ByRef p_txt As TextBox)
  p_txt.Text = "GNU GENERAL PUBLIC LICENSE" & vbCrLf & _
    "Version 2, June 1991 " & vbCrLf & _
    "" & vbCrLf & _
    "Copyright (C) 1989, 1991 Free Software Foundation, Inc.  " & vbCrLf & _
    "59 Temple Place - Suite 330, Boston, MA  02111-1307, USA" & vbCrLf & _
    "" & vbCrLf & _
    "Everyone is permitted to copy and distribute verbatim copies" & vbCrLf & _
    "of this license document, but changing it is not allowed." & vbCrLf & _
    "" & vbCrLf & _
    "Preamble" & vbCrLf & _
    "The licenses for most software are designed to take away your freedom to share and change it. By contrast, the GNU General Public License is intended to guarantee your freedom to share and change free software--to make sure the software is free for all its users. This General Public License applies to most of the Free Software Foundation's software and to any other program whose authors commit to using it. (Some other Free Software Foundation software is covered by the GNU Library General Public License instead.) You can apply it to your programs, too. " & vbCrLf & _
    "" & vbCrLf & _
    "When we speak of free software, we are referring to freedom, not price. Our General Public Licenses are designed to make sure that you have the freedom to distribute copies of free software (and charge for this service if you wish), that you receive source code or can get it if you want it, that you can change the software or use pieces of it in new free programs; and that you know you can do these things. " & vbCrLf & _
    "" & vbCrLf & _
    "To protect your rights, we need to make restrictions that forbid anyone to deny you these rights or to ask you to surrender the rights. These restrictions translate to certain responsibilities for you if you distribute copies of the software, or if you modify it. " & vbCrLf & _
    "" & vbCrLf & _
    "For example, if you distribute copies of such a program, whether gratis or for a fee, you must give the recipients all the rights that you have. You must make sure that they, too, receive or can get the source code. And you must show them these terms so they know their rights. " & vbCrLf & _
    "" & vbCrLf
  p_txt.Text = p_txt.Text & "TERMS AND CONDITIONS FOR COPYING, DISTRIBUTION AND MODIFICATION" & vbCrLf & _
    "0. This License applies to any program or other work which contains a notice placed by the copyright holder saying it may be distributed under the terms of this General Public License. The ""Program"", below, refers to any such program or work, and a ""work based on the Program"" means either the Program or any derivative work under copyright law: that is to say, a work containing the Program or a portion of it, either verbatim or with modifications and/or translated into another language. (Hereinafter, translation is included without limitation in the term ""modification"".) Each licensee is addressed as ""you"". " & vbCrLf & _
    "" & vbCrLf & _
    "Activities other than copying, distribution and modification are not covered by this License; they are outside its scope. The act of running the Program is not restricted, and the output from the Program is covered only if its contents constitute a work based on the Program (independent of having been made by running the Program). Whether that is true depends on what the Program does. " & vbCrLf & _
    "" & vbCrLf & _
    "1. You may copy and distribute verbatim copies of the Program's source code as you receive it, in any medium, provided that you conspicuously and appropriately publish on each copy an appropriate copyright notice and disclaimer of warranty; keep intact all the notices that refer to this License and to the absence of any warranty; and give any other recipients of the Program a copy of this License along with the Program. " & vbCrLf & _
    "" & vbCrLf & _
    "You may charge a fee for the physical act of transferring a copy, and you may at your option offer warranty protection in exchange for a fee. " & vbCrLf & _
    "" & vbCrLf & _
    "2. You may modify your copy or copies of the Program or any portion of it, thus forming a work based on the Program, and copy and distribute such modifications or work under the terms of Section 1 above, provided that you also meet all of these conditions: " & vbCrLf & _
    "" & vbCrLf & _
    "" & vbCrLf & _
    "a) You must cause the modified files to carry prominent notices stating that you changed the files and the date of any change. " & vbCrLf & _
    "" & vbCrLf & _
    "b) You must cause any work that you distribute or publish, that in whole or in part contains or is derived from the Program or any part thereof, to be licensed as a whole at no charge to all third parties under the terms of this License. " & vbCrLf & _
    "" & vbCrLf & _
    "c) If the modified program normally reads commands interactively when run, you must cause it, when started running for such interactive use in the most ordinary way, to print or display an announcement including an appropriate copyright notice and a notice that there is no warranty (or else, saying that you provide a warranty) and that users may redistribute the program under these conditions, and telling the user how to view a copy of this License. (Exception: if the Program itself is interactive but does not normally print such an announcement, your work based on the Program is not required to print an announcement.) " & vbCrLf & _
    "These requirements apply to the modified work as a whole. If identifiable sections of that work are not derived from the Program, and can be reasonably considered independent and separate works in themselves, then this License, and its terms, do not apply to those sections when you distribute them as separate works. But when you distribute the same sections as part of a whole which is a work based on the Program, the distribution of the whole must be on the terms of this License, whose permissions for other licensees extend to the entire whole, and thus to each and every part regardless of who wrote it. " & vbCrLf
  p_txt.Text = p_txt.Text & "Thus, it is not the intent of this section to claim rights or contest your rights to work written entirely by you; rather, the intent is to exercise the right to control the distribution of derivative or collective works based on the Program. " & vbCrLf & _
    "" & vbCrLf & _
    "In addition, mere aggregation of another work not based on the Program with the Program (or with a work based on the Program) on a volume of a storage or distribution medium does not bring the other work under the scope of this License. " & vbCrLf & _
    "" & vbCrLf & _
    "3. You may copy and distribute the Program (or a work based on it, under Section 2) in object code or executable form under the terms of Sections 1 and 2 above provided that you also do one of the following: " & vbCrLf & _
    "" & vbCrLf & _
    "a) Accompany it with the complete corresponding machine-readable source code, which must be distributed under the terms of Sections 1 and 2 above on a medium customarily used for software interchange; or, " & vbCrLf & _
    "" & vbCrLf & _
    "b) Accompany it with a written offer, valid for at least three years, to give any third party, for a charge no more than your cost of physically performing source distribution, a complete machine-readable copy of the corresponding source code, to be distributed under the terms of Sections 1 and 2 above on a medium customarily used for software interchange; or, " & vbCrLf & _
    "" & vbCrLf & _
    "c) Accompany it with the information you received as to the offer to distribute corresponding source code. (This alternative is allowed only for noncommercial distribution and only if you received the program in object code or executable form with such an offer, in accord with Subsection b above.) " & vbCrLf & _
    "The source code for a work means the preferred form of the work for making modifications to it. For an executable work, complete source code means all the source code for all modules it contains, plus any associated interface definition files, plus the scripts used to control compilation and installation of the executable. However, as a special exception, the source code distributed need not include anything that is normally distributed (in either source or binary form) with the major components (compiler, kernel, and so on) of the operating system on which the executable runs, unless that component itself accompanies the executable. " & vbCrLf & _
    "If distribution of executable or object code is made by offering access to copy from a designated place, then offering equivalent access to copy the source code from the same place counts as distribution of the source code, even though third parties are not compelled to copy the source along with the object code. " & vbCrLf & _
    "" & vbCrLf & _
    "4. You may not copy, modify, sublicense, or distribute the Program except as expressly provided under this License. Any attempt otherwise to copy, modify, sublicense or distribute the Program is void, and will automatically terminate your rights under this License. However, parties who have received copies, or rights, from you under this License will not have their licenses terminated so long as such parties remain in full compliance. " & vbCrLf & _
    "" & vbCrLf & _
    "5. You are not required to accept this License, since you have not signed it. However, nothing else grants you permission to modify or distribute the Program or its derivative works. These actions are prohibited by law if you do not accept this License. Therefore, by modifying or distributing the Program (or any work based on the Program), you indicate your acceptance of this License to do so, and all its terms and conditions for copying, distributing or modifying the Program or works based on it. " & vbCrLf & _
    "" & vbCrLf
  p_txt.Text = p_txt.Text & "6. Each time you redistribute the Program (or any work based on the Program), the recipient automatically receives a license from the original licensor to copy, distribute or modify the Program subject to these terms and conditions. You may not impose any further restrictions on the recipients' exercise of the rights granted herein. You are not responsible for enforcing compliance by third parties to this License. " & vbCrLf & _
    "" & vbCrLf & _
    "7. If, as a consequence of a court judgment or allegation of patent infringement or for any other reason (not limited to patent issues), conditions are imposed on you (whether by court order, agreement or otherwise) that contradict the conditions of this License, they do not excuse you from the conditions of this License. If you cannot distribute so as to satisfy simultaneously your obligations under this License and any other pertinent obligations, then as a consequence you may not distribute the Program at all. For example, if a patent license would not permit royalty-free redistribution of the Program by all those who receive copies directly or indirectly through you, then the only way you could satisfy both it and this License would be to refrain entirely from distribution of the Program. " & vbCrLf & _
    "" & vbCrLf & _
    "If any portion of this section is held invalid or unenforceable under any particular circumstance, the balance of the section is intended to apply and the section as a whole is intended to apply in other circumstances. " & vbCrLf & _
    "" & vbCrLf & _
    "It is not the purpose of this section to induce you to infringe any patents or other property right claims or to contest validity of any such claims; this section has the sole purpose of protecting the integrity of the free software distribution system, which is implemented by public license practices. Many people have made generous contributions to the wide range of software distributed through that system in reliance on consistent application of that system; it is up to the author/donor to decide if he or she is willing to distribute software through any other system and a licensee cannot impose that choice. " & vbCrLf & _
    "" & vbCrLf & _
    "This section is intended to make thoroughly clear what is believed to be a consequence of the rest of this License. " & vbCrLf & _
    "" & vbCrLf & _
    "8. If the distribution and/or use of the Program is restricted in certain countries either by patents or by copyrighted interfaces, the original copyright holder who places the Program under this License may add an explicit geographical distribution limitation excluding those countries, so that distribution is permitted only in or among countries not thus excluded. In such case, this License incorporates the limitation as if written in the body of this License. " & vbCrLf & _
    "" & vbCrLf & _
    "9. The Free Software Foundation may publish revised and/or new versions of the General Public License from time to time. Such new versions will be similar in spirit to the present version, but may differ in detail to address new problems or concerns. " & vbCrLf & _
    "" & vbCrLf & _
    "Each version is given a distinguishing version number. If the Program specifies a version number of this License which applies to it and ""any later version"", you have the option of following the terms and conditions either of that version or of any later version published by the Free Software Foundation. If the Program does not specify a version number of this License, you may choose any version ever published by the Free Software Foundation. " & vbCrLf & _
    "" & vbCrLf & _
    "10. If you wish to incorporate parts of the Program into other free programs whose distribution conditions are different, write to the author to ask for permission. For software which is copyrighted by the Free Software Foundation, write to the Free Software Foundation; we sometimes make exceptions for this. Our decision will be guided by the two goals of preserving the free status of all derivatives of our free software and of promoting the sharing and reuse of software generally. " & vbCrLf & _
    "" & vbCrLf & _
    "NO WARRANTY" & vbCrLf & _
    "" & vbCrLf & _
    "11. BECAUSE THE PROGRAM IS LICENSED FREE OF CHARGE, THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY APPLICABLE LAW. EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM ""AS IS"" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM IS WITH YOU. SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF ALL NECESSARY SERVICING, REPAIR OR CORRECTION. " & vbCrLf & _
    "" & vbCrLf & _
    "12. IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MAY MODIFY AND/OR REDISTRIBUTE THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES, INCLUDING ANY GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY YOU OR THIRD PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER PROGRAMS), EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. " & vbCrLf & _
    ""
End Sub

'-------------------------------------------------------------------------------
' BEHAVIOR    : project dependent
' PROCEDURE   : TurnOn(p_TurnOn, p_who)
' DESCRIPTION : This function will set a control cosmetic properties
'               depending on the mouse behavior: OnMouseOver, OnMouseOff.
'               You should modify the control's cosmetics...
' NOTE        : This function is being called by the global "OnMouseOver"
'               that can be found in the LIB_Controls library module.
' PARAMETERS  : - p_TurnOn: Boolean     To turn on or off cosmetics
'               - p_who: Object         The control's pointer.
' EXAMPLE     : TurnOn False, Label1
'
'  p_who.BorderStyle = IIf(p_TurnOn, vbFixedSingle, vbBSNone)
'  p_who.BackColor = IIf(p_TurnOn, vbGrayText, &HE0E0E0)
'  p_who.FontBold = p_TurnOn
'  p_who.ForeColor = IIf(p_TurnOn, vbBlue, vbBlack)
'  p_who.FontUnderline = p_TurnOn
'  Screen.MousePointer = IIf(p_TurnOn, vbCrosshair, vbDefault)
'
' Copyright © 2001 Daniele Giabbai <DGiabbai@hotpop.com>
'-------------------------------------------------------------------------------
Public Sub TurnOn(p_TurnOn As Boolean, p_who As Object)
  On Error Resume Next
  If p_who = lblProgramVersion Then Exit Sub
  If Not p_who Is txtLicense Then
    Select Case TypeName(p_who)
      Case "Image"
        p_who.BorderStyle = IIf(p_TurnOn, vbFixedSingle, vbBSNone)
        Screen.MousePointer = IIf(p_TurnOn, vbCrosshair, vbDefault)
      Case "Label"
        p_who.ForeColor = IIf(p_TurnOn, vbBlue, vbBlack)
        p_who.Font.Underline = p_TurnOn
        Screen.MousePointer = IIf(p_TurnOn, vbCustom, vbDefault)
        Screen.MouseIcon = IIf(p_TurnOn, imgMouseOver.Picture, vbDefault)
      Case "TextBox"
        p_who.ForeColor = IIf(p_TurnOn, vbBlue, vbBlack)
      Case "CommandButton"
        p_who.FontBold = p_TurnOn
    End Select
  End If
End Sub
'
'************************************************************************
'************************************************************************
'

Private Sub cmdOk_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  OnMouseOver Me, cmdOk
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdOk_Click
  If KeyAscii = vbKeyEscape Then cmdOk_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  OnMouseOver Me
End Sub

Private Sub Form_Resize()
  cmdOk.Top = Me.ScaleHeight - cmdOk.Height - 100
  cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
  
  frameInfo.Width = Me.ScaleWidth - 300
  frameInfo.Left = (Me.ScaleWidth - frameInfo.Width) / 2
  frameInfo.Height = cmdOk.Top - frameInfo.Top - 100
  
  txtInfo.Top = 200
  txtInfo.Left = 100
  txtInfo.Height = 750
  txtInfo.Width = frameInfo.Width - txtInfo.Left * 2
  
  txtLicense.Top = txtInfo.Top + txtInfo.Height + 50
  txtLicense.Left = 100
  txtLicense.Width = frameInfo.Width - txtInfo.Left * 2
  txtLicense.Height = frameInfo.Height - txtInfo.Top * 1.5 - txtInfo.Height - 50
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  'OnMouseOver Me, Image1
End Sub

Private Sub lblProgramAuthor_Click()
  Exec lblProgramAuthor.ToolTipText
End Sub

Private Sub lblProgramAuthor_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  OnMouseOver Me, lblProgramAuthor
End Sub

Private Sub lblProgramName_Click()
  Exec lblProgramName.ToolTipText
End Sub

Private Sub lblProgramName_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  OnMouseOver Me, lblProgramName
End Sub

Private Sub lblProgramVersion_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  OnMouseOver Me, lblProgramVersion
End Sub

Private Sub txtInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  OnMouseOver Me, txtInfo
End Sub
