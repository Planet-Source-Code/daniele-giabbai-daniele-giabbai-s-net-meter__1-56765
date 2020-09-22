Attribute VB_Name = "LIB_Controls"
Option Explicit

'-------------------------------------------------------------------------------
' PROCEDURE   : AddElement
' DESCRIPTION : This procedure adds elements to a listbox or combobox allowing
'   to not add duplicate elements. Works fine with small lists...
' EXAMPLE     :
'   Sub TestAddTwoElements()
'     AddElement pListObj, "test element"
'     AddElement pListObj, "test element"
'   End Sub
' Copyright © 2003 by Daniele Giabbai
'-------------------------------------------------------------------------------
Public Sub AddElement(pListObj As Object, pItem As String, Optional NoDup As Boolean = True)
  If NoDup Then
    Dim i As Long
    'Debug.Print pListObj.ListCount
    For i = 0 To pListObj.ListCount
      'Debug.Print i & " - " & pListObj.List(i)
      If pListObj.List(i) = pItem Then Exit Sub
    Next
  End If
  pListObj.AddItem pItem
End Sub

Sub AssignCaption(pCtrl As Control, pWhat As String)
  If TypeOf pCtrl Is Label Then
    pCtrl.Caption = pWhat
  ElseIf TypeOf pCtrl Is TextBox Then
    pCtrl.Text = pWhat
  End If
End Sub

Sub Append(ByRef p_txt As Object, Optional p_what As String)
  p_txt = p_txt & p_what
End Sub

'Function: Completes a word by searching through a specified listbox
'will skip as many matches as the number you type in for Skip
Public Function AutoCompleteOld(Word As String, List As ListBox, Optional ByVal Skip As Integer = 0) As String
  Dim i As Integer
  Dim j As Integer

  For i = 1 To Len(Word)
    For j = 0 To List.ListCount - 1
      If UCase(Left(Word, 1)) = UCase(List.List(j)) Then
        If Skip > 0 Then
          Skip = Skip - 1
        Else
          AutoCompleteOld = List.List(i)
          Exit Function
        End If
      End If
    Next
  Next
End Function

Sub AutoComplete(txt As TextBox, pWord As String, pList As ListBox)
  Dim i As Integer
End Sub

Public Sub CheckBox_SetValue(ByRef p_Source As CheckBox, p_value As Boolean)
  If p_value Then
    p_Source.Value = vbChecked
  Else
    p_Source.Value = vbUnchecked
  End If
  CheckBox_SetToolTip p_Source
End Sub

Public Sub CheckBox_SetToolTip(ByRef p_Source As CheckBox)
  If p_Source.Value = vbChecked Then
    p_Source.ToolTipText = p_Source.Tag & " ON"
  Else
    p_Source.ToolTipText = p_Source.Tag & " OFF"
  End If
End Sub

' copies only non empty rows of a combo
Public Sub ComboBox_Copy(ByRef Source As ComboBox, ByRef Destination As ComboBox)
  Dim i, j, Count As Integer
  
  Count = Source.ListCount - 1
  For i = 0 To Count
    If Source.List(i) <> "" Then Destination.AddItem Source.List(i)
  Next i
End Sub

' copies all rows of a combo
Public Sub ComboBox_Mirror(ByRef Source As ComboBox, ByRef Destination As ComboBox)
  Dim i, Count As Integer
  
  Count = Source.ListCount - 1
  For i = 0 To Count
    Destination.List(i) = Source.List(i)
  Next i
End Sub

Public Function ComboBox_Search(ByRef Source As ComboBox, p_search As String) As Long
  Dim i, Count As Integer
  
  ComboBox_Search = -1
  Count = Source.ListCount - 1
  For i = 0 To Count
    If p_search = Source.List(i) Then
      ComboBox_Search = i
      Exit Function
    End If
  Next i
End Function

Public Sub EnableFrameControls(frameFrame As Frame, Optional enable As Boolean = True)
  Dim obj As Object
  
  For Each obj In frameFrame.Parent.Controls
    If obj.Container Is frameFrame Then
      On Error Resume Next
'      If TypeOf Obj Is CommandButton Then Obj.Visible = True
      obj.Enabled = enable
    End If
  Next
  frameFrame.Enabled = enable
End Sub

'-------------------------------------------------------------------------------
' BEHAVIOR    : project independent
' PROCEDURE   : LoadControls
' DESCRIPTION : This procedure loads controls in a control array.
' PARAMETERS  : - ReferenceControl         The control in the control array
'               - iNumberOfControls        The number of controls to be loaded
'               - [pTiled]                 Show controls, one beside the other
'               - [iTiled_MaxPerLine]      Max controls per line
'               - [iTiled_LeftMargin]      Margin to the previous control to the left
'               - [iTiled_TopMargin]       Margin to the previous control to the top
'               - [ShowCaption]            Assign a caption to the control, equal to the control index
'               - [pBackColor]             Control's back color
'               - [pForeColor]             Control's fore color
'               - [pEnabled]               Enable control
'
' REMARKS     : If pTiled = False, then the following parameters will not have effect:
'    [iTiled_MaxPerLine], [iTiled_LeftMargin], [iTiled_TopMargin],
'    [ShowCaption], [pBackColor]
'
' NOTE        : When creating the control array, draw the desired control in the
'   form to the *proper* position, as the loaded controls will follow it.
'
' EXAMPLE     :
'   Load 30 timer controls, non-tiled:
'
'     Call LoadControls(Timer1, 30)
'
'   Load 255 label controls, tiled, with 20 labels per row
'
'     Call LoadControls(Label1, 255, True, 20)
'
' Copyright © 2001 Daniele Giabbai <DGiabbai@hotpop.com>
'-------------------------------------------------------------------------------
Sub LoadControls( _
      ReferenceControl As Object _
      , iNumberOfControls As Integer _
      , Optional pTiled As Boolean = False _
      , Optional iTiled_MaxPerLine As Integer = 10 _
      , Optional iTiled_LeftMargin As Integer = 15 _
      , Optional iTiled_TopMargin As Integer = 15 _
      , Optional ShowCaption As Boolean = True _
      , Optional pBackColor As OLE_COLOR = &HFFFFFF _
      , Optional pForeColor As OLE_COLOR = &H0& _
      , Optional pEnabled As Boolean = True _
      )
  
  Dim i As Integer, k As Integer
  Dim lTOp As Long, lLeft As Long, lWidth As Long, lHeight As Long
  
  i = ReferenceControl.LBound
  If pTiled Then
    lTOp = ReferenceControl(i).Top
    lLeft = ReferenceControl(i).Left
    lWidth = ReferenceControl(i).Width
    lHeight = ReferenceControl(i).Height
    ReferenceControl(i).BackColor = pBackColor
    ReferenceControl(i).ForeColor = pForeColor
  End If
  AssignCaption ReferenceControl(i), IIf(ShowCaption, "1", "")
  i = i + 1
  
  For k = 2 To iNumberOfControls
    Load ReferenceControl(i)
    ReferenceControl(i).Enabled = pEnabled
    If pTiled Then
      ReferenceControl(i).BackColor = pBackColor
      ReferenceControl(i).ForeColor = pForeColor
      ReferenceControl(i).Top = lTOp
      'ReferenceControl(i).Width = lWidth
      ReferenceControl(i).Left = lLeft + (lWidth + iTiled_LeftMargin) * ((k - 1) Mod iTiled_MaxPerLine)
      ReferenceControl(i).Visible = True
      AssignCaption ReferenceControl(i), IIf(ShowCaption, CStr(k), "")
      If (k Mod iTiled_MaxPerLine) = 0 Then
        lTOp = lTOp + lHeight + iTiled_TopMargin
      End If
    End If
    i = i + 1
  Next k
End Sub

'-------------------------------------------------------------------------------
'* BETA!
'* Che succede se il controllo successivo ha la proprietà tabstop=false?
'-------------------------------------------------------------------------------
Public Sub NextFieldFocus(xForm As Form)
  On Error Resume Next
  Dim Ctl As Control
  Dim m_tabIndex As Long
  Dim SearchCtl As Control
  
  ' search for the control that has the greater tabindex value and
  ' that has tabstop set to true
  m_tabIndex = xForm.ActiveControl.TabIndex
  For Each Ctl In xForm.Controls
    If Ctl.TabIndex > m_tabIndex And Ctl.TabStop Then Set SearchCtl = Ctl
  Next
  If Not SearchCtl Is Nothing Then SearchCtl.SetFocus
  
End Sub

'*************************************************************
'* Formats text in a textbox, label or string handling correctly
'* the change event.
'
'Private Sub Text1_Change()
'  If OnChange_Format(Text1, "##,##0.00",CurrencyDecimalSeparator) Then CalculateChange
'End Sub
'
Public Function OnChange_Format(ByRef obj As Variant, frmt As String, Optional charToChangeBehavior As Variant) As Boolean
  Dim s, txt As String
  
  ' TextBox
  If TypeOf obj Is TextBox Then
    s = Trim$(obj.Text)
    If Len(s) >= 1 Then
      If Not IsNumeric(Left$(s, 1)) Then
        obj.Text = Right$(s, Len(s) - 1)
        Exit Function
      End If
    End If
'    If Not IsMissing(charToChangeBehavior) Then
'      If InStr(1, s, charToChangeBehavior) = 0 Then
'        If Len(s) >= 3 Then
'          'Debug.Print "Prima:" & s
'          s = Left$(s, Len(s) - 2) & charToChangeBehavior & Right$(s, 2)
'          'Debug.Print "Dopo :" & s
'        End If
'      End If
'    End If
    s = Format(s, frmt)
    If s <> obj.Text Then
      Dim i, cursPos, decimalPos As Long
      Dim normalBehaviour As Boolean
      
      cursPos = obj.SelStart ' remember cursor position
      i = Len(obj.Text) - cursPos ' remember how many chars are to the right of cursor
      
      If Not IsMissing(charToChangeBehavior) Then
        decimalPos = InStr(1, s, charToChangeBehavior)
        normalBehaviour = (obj.SelStart > decimalPos)
      End If
      
      obj.Text = s ' this assignement re-fires the "change" event
      If Len(s) = 4 Then ' case: text="1,00"
        obj.SelStart = 1 ' set cursor position after first char
      ElseIf Not normalBehaviour Then
        If i < Len(s) Then
          obj.SelStart = Len(s) - i
        Else
          obj.SelStart = 0
        End If
      Else
        obj.SelStart = cursPos
      End If
    Else
      OnChange_Format = True
    End If
  
  ' Label
  ElseIf TypeOf obj Is Label Then
    s = Format(Trim(obj.Caption), frmt)
    If s <> obj.Caption Then
      obj.Caption = s ' this assignement re-fires the "change" event
    Else
      OnChange_Format = True
    End If
  
  ' String
  ElseIf TypeName(obj) = "String" Then
    obj = Format(Trim(obj), frmt)
  End If
End Function

'*************************************************************
'* Limit the size of a textbox
'
'Private Sub Text1_Change()
'  OnChange_LimitSize Text1, 50
'End Sub
'
Public Sub OnChange_LimitSize(ByRef txt As TextBox, size As Long)
  txt.Text = Left$(txt.Text, size)
End Sub

Public Sub OnGotFocus(ctrl As Control, s As String, Optional color As Long = &H80000006)
  If ctrl.Text = s Then ctrl.Text = ""
  ctrl.ForeColor = color
End Sub

Public Sub OnLostFocus(ctrl As Control, s As String, Optional color As Long = &H80000004)
  If Len(ctrl.Text) = 0 Then
    ctrl.Text = s
    ctrl.ForeColor = color
  End If
End Sub

'-------------------------------------------------------------------------------
' BEHAVIOR    : project independent
' FUNCTION    : OnKeyPress_AllowKeys
' DESCRIPTION : Enable only few characters to be pressed.
' PARAMETERS  : - KeyAscii As Integer         The character searched.
'               - pAcceptedChars As String    Accepted chars.
' EXAMPLE     :
'    Private Sub txtNumber_KeyPress(KeyAscii As Integer)
'      KeyAscii = OnKeyPress_AllowKeys(KeyAscii, "-0123456789.,")
'    End Sub
'
' Copyright © 2001 Daniele Giabbai <DGiabbai@hotpop.com>
'-------------------------------------------------------------------------------
Public Function OnKeyPress_AllowKeys( _
      KeyAscii As Integer _
      , Optional pAcceptedChars As String = "." _
      ) As Integer
  
  If KeyAscii = vbKeyBack Then
    OnKeyPress_AllowKeys = KeyAscii
    Exit Function
  ElseIf InStr(1, pAcceptedChars, Chr$(KeyAscii)) <> 0 Then
    OnKeyPress_AllowKeys = KeyAscii
    Exit Function
  End If
  OnKeyPress_AllowKeys = 0
End Function

'-------------------------------------------------------------------------------
' BEHAVIOR    : project independent
' FUNCTION    : OnKeyPress_AllowNumericKeysOnly
' DESCRIPTION : Enable only numeric characters to be pressed.
' PARAMETERS  : - KeyAscii As Integer         The character searched.
' EXAMPLE     :
'    Private Sub txtNumber_KeyPress(KeyAscii As Integer)
'      KeyAscii = OnKeyPress_AllowNumericKeysOnly(KeyAscii)
'    End Sub
'
' Copyright © 2001 Daniele Giabbai <DGiabbai@hotpop.com>
'-------------------------------------------------------------------------------
Public Function OnKeyPress_AllowNumericKeysOnly(KeyAscii As Integer) As Integer
  If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or _
      KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd _
    Then
    OnKeyPress_AllowNumericKeysOnly = KeyAscii
    Exit Function
  ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    OnKeyPress_AllowNumericKeysOnly = KeyAscii
    Exit Function
  ElseIf KeyAscii >= Asc(".") Or KeyAscii <= Asc(",") Or KeyAscii <= Asc("-") Then
    OnKeyPress_AllowNumericKeysOnly = KeyAscii
    Exit Function
  End If
  OnKeyPress_AllowNumericKeysOnly = 0
End Function

Public Function OnKeyPress_Format(KeyAscii As Integer, ByRef obj As TextBox) As Integer
  Dim sep As String
  If KeyAscii = vbKeyBack Then
    If obj.SelStart > 0 Then
      sep = Mid(obj.Text, obj.SelStart, 1)
      If sep = "." Or sep = "," Then
        obj.SelStart = obj.SelStart - 1
        obj.SelLength = 0
        OnKeyPress_Format = 0
        Exit Function
      End If
    End If
  ElseIf KeyAscii = vbKeyDelete Then
    If obj.SelStart < Len(obj.Text) Then
      sep = Mid(obj.Text, obj.SelStart + 1, 1)
      If sep = "." Or sep = "," Then
        obj.SelStart = obj.SelStart + 1
        obj.SelLength = 0
        OnKeyPress_Format = 0
        Exit Function
      End If
    End If
  End If
  OnKeyPress_Format = KeyAscii
End Function

'-------------------------------------------------------------------------------
' BEHAVIOR    : project independent
' FUNCTION    : OnMouseOver form, [p_who]
' DESCRIPTION : This function controls the mouse behavior: OnMouseOver, OnMouseOff.
'               You should modify the control's cosmetics in the function "TurnOn"
'               With no parameters it will turn OFF cosmetics from a control (behavior
'               like OnMouseOff). You should call it with no parameters from the
'               "Form_MouseMove" subroutine.
'               With one parameter, the control's pointer, it will turn the control's
'               cosmetics ON (behavior like OnMouseOver). You should call it from the
'               control's "_MouseMove" method.
' PARAMETERS  : - frm: Form          The form in which is defined the function "TurnOn"
'               - [p_who]: Object    The control's pointer.
' EXAMPLE     :
'
'    Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      OnMouseOver Me, Label1 ' to turn cosmetics on
'    End Sub
'
'    Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      OnMouseOver Me, Label2 ' to turn cosmetics on
'    End Sub
'
'    Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      ' To turn cosmetics off from *ANY* control for wich cosmetics were turned on
'      OnMouseOver Me
'    End Sub
'
'    Public Sub TurnOn(p_TurnOn As Boolean, p_who As Object)
'      On Error Resume Next
'      Select Case TypeName(p_who)
'        Case "Label"
'          p_who.FontUnderline = p_TurnOn
'        Case "TextBox"
'          If p_TurnOn Then
'            Screen.MousePointer = vbCrosshair
'          Else
'            Screen.MousePointer = vbDefault
'          End If
'      End Select
'    End Sub
'
' Copyright © 2001 Daniele Giabbai <DGiabbai@hotpop.com>
'-------------------------------------------------------------------------------
Public Sub OnMouseOver(frm As Form, Optional p_who As Object)
  Static previous_who As Object
  
  If Not p_who Is Nothing Then ' activating new one
    If Not previous_who Is Nothing Then
      If previous_who = p_who Then Exit Sub ' control is already active
      frm.TurnOn False, previous_who ' disactivate previous
    End If
    frm.TurnOn True, p_who ' activate new
    Set previous_who = p_who ' remember last one
  ElseIf Not previous_who Is Nothing Then ' disactivate last one
    frm.TurnOn False, previous_who ' disactivate previous
    Set previous_who = Nothing ' do not disactivate it again and again and ...
  End If
End Sub

'-------------------------------------------------------------------------------
' PROCEDURE   : RemoveDuplicates
' DESCRIPTION : This procedure removes duplicate items from a listbox or combobox
' EXAMPLE     :
' Copyright © 2003 by Daniele Giabbai
'-------------------------------------------------------------------------------
Public Sub RemoveDuplicates(pListObj As Object)
  If pListObj.Sorted Then
    Dim i As Long
    
    For i = 0 To pListObj.ListCount - 1
    Next
  Else
  End If
End Sub

'-------------------------------------------------------------------------------
'* Selects the text in a control
'
'Private Sub Text1_GotFocus()
'  Sel Me.ActiveControl
'End Sub
'
'-------------------------------------------------------------------------------
Public Sub Sel(Optional ctrl As Control = Nothing)
  On Error Resume Next
  If ctrl Is Nothing Then Set ctrl = Screen.ActiveForm.ActiveControl
  If TypeOf ctrl Is TextBox Then
    ctrl.SelStart = 0
    ctrl.SelLength = Len(ctrl.Text)
  End If
End Sub

Sub UnloadControls(ReferenceControl As Object)
  On Error Resume Next
  Dim i As Integer, k As Integer
  
  k = ReferenceControl.LBound + 1
  For i = ReferenceControl.UBound To k Step -1
    Unload ReferenceControl(i)
  Next i
  On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' BEHAVIOR    : project independent
' FUNCTION    : WordSearch
' DESCRIPTION : This function finds the first word in an ordered list which is
'               most similar to the one searched.
' PARAMETERS  : - pWord As String   The word searched.
'               - pList As ListBox  The object of the search.
' Copyright © 2001 Daniele Giabbai <DGiabbai@hotpop.com>
'-------------------------------------------------------------------------------
Public Function WordSearch(pWord As String, pList As ListBox) As Integer
  Dim letter As Integer
  Dim j As Integer
  Dim i As Integer
  
  i = 0
  j = 0
  If pList.ListCount <> 0 Then
    letter = 1
    Do While letter <= Len(pWord)
      Do While LCase$(Left$(pWord, letter)) <> LCase(Left$(pList.List(j), letter))
        j = j + 1
        If j = pList.ListCount Then
          WordSearch = IIf((letter = 1), 0, i)
          Exit Function
        End If
      Loop
      letter = letter + 1
      i = j
    Loop
  End If
  'Debug.Print p_List.List(i)
  WordSearch = i
End Function


