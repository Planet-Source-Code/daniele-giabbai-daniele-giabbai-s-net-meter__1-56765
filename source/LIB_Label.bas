Attribute VB_Name = "LIB_Label"
Option Explicit

Sub LabelFontResize(lbl As Label, maxW As Long)
  If Not lbl.Visible Then Exit Sub
  lbl.AutoSize = True
  If lbl.Width < maxW Then
    Call FontIncrease(lbl, maxW)
    If lbl.Width > maxW Then
      lbl.FontSize = lbl.FontSize - 1
    End If
  ElseIf lbl.Width > maxW Then
    Call FontReduce(lbl, maxW)
  End If
  lbl.AutoSize = False
End Sub

Sub FontIncrease(lbl As Label, maxW As Long)
  If lbl.Width < maxW Then
    If lbl.FontSize < 10 Then
      lbl.FontSize = lbl.FontSize + 1
      Call FontIncrease(lbl, maxW)
    Else
      lbl.FontSize = 10
    End If
  End If
End Sub

Sub FontReduce(lbl As Label, maxW As Long)
  If lbl.Width > maxW Then
    If lbl.FontSize > 6 Then
      lbl.FontSize = lbl.FontSize - 1
      Call FontReduce(lbl, maxW)
    Else
      lbl.FontSize = 6
    End If
  End If
End Sub
