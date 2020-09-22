Attribute VB_Name = "modGraph"
Option Explicit

Sub DrawUsage(picGraph As PictureBox, lDownloadBytes As Long, lUploadBytes As Long, Optional RedrawOnly As Boolean = False)
  On Error GoTo ErrH
  '
  ' Variables setted once
  '
  Static UboundDLValues As Long
  Static inited As Boolean
  
  '
  ' Variables updated at each call
  '
  Static DLValues() As Long
  Static ULValues() As Long
  Static lCurrentIndex As Long
  Static lMaxValue As Long
  
  If lDownloadBytes > 10000000 Then Exit Sub ' 10Mb transfer rate? Mhmm...
  If lUploadBytes > 10000000 Then Exit Sub   ' 10Mb transfer rate? Mhmm...
  If Not inited Then
    ReDim DLValues(0 To (Screen.Width \ Screen.TwipsPerPixelX))
    ReDim ULValues(0 To Screen.Width \ Screen.TwipsPerPixelX)
    UboundDLValues = UBound(DLValues)
    lCurrentIndex = UBound(DLValues)
    ' Initialization of the lMaxValue to a positive value
    ' prevents from checking it > 0 in the main drawing cycle.
    ' as it is used as a denominator in the proportion below.
    lMaxValue = picGraph.ScaleHeight
    inited = True
  Else
    ' Here we fill values backwards, so that in the main drawing cycle below we can
    ' use a more efficient "(lIndex + 1) Mod UboundDLValues" operation
    ' than decrementing and checking by "if" statement like here...
    lCurrentIndex = (lCurrentIndex - 1)
    If lCurrentIndex < 0 Then lCurrentIndex = UboundDLValues
  End If
  
  '
  ' Update values
  '
  DLValues(lCurrentIndex) = lDownloadBytes
  ULValues(lCurrentIndex) = lUploadBytes
  If lMaxValue < lDownloadBytes Then lMaxValue = lDownloadBytes
  If lMaxValue < lUploadBytes Then lMaxValue = lUploadBytes
  If lMaxValue = 0 Then Exit Sub
  
  '
  ' Redraw
  '
  Dim X As Long
  Dim lIndex As Long
  Dim lDrawValue As Long
  Dim YMax As Long
  Dim lMaxPartial As Long
  
  picGraph.Cls
  YMax = picGraph.ScaleHeight   ' Get the graph max heigth
  lIndex = lCurrentIndex
  Debug.Print "DL: " & DLValues(lIndex)
  Debug.Print "UL: " & ULValues(lIndex)
  For X = picGraph.ScaleWidth To 0 Step -1
    
    ' We need to scale the values to the available window dimension.
    ' And that scaling is performed through this simple proportion:
    '   lDrawValue : DLValues(X) = YMax : lMaxValue
    ' which is used below.
    If DLValues(lIndex) = ULValues(lIndex) Then
      
      '
      ' Case 1
      '
1      lDrawValue = ((DLValues(lIndex) * YMax) / lMaxValue)
      If lDrawValue > 0 Then
        picGraph.Line (X, YMax)-(X, YMax - lDrawValue), vbRed
        If lMaxPartial < DLValues(lIndex) Then lMaxPartial = DLValues(lIndex)
      End If
    
    ElseIf DLValues(lIndex) > ULValues(lIndex) Then
      
      '
      ' Case 2
      '
2      lDrawValue = ((DLValues(lIndex) * YMax) / lMaxValue)
      If lDrawValue > 0 Then
        picGraph.Line (X, YMax)-(X, YMax - lDrawValue), vbRed
        If lMaxPartial < DLValues(lIndex) Then lMaxPartial = DLValues(lIndex)
      End If
      
3      lDrawValue = ((ULValues(lIndex) * YMax) / lMaxValue)
      If lDrawValue > 0 Then
        picGraph.Line (X, YMax)-(X, YMax - lDrawValue), vbYellow
        If lMaxPartial < ULValues(lIndex) Then lMaxPartial = ULValues(lIndex)
      End If
      
    Else
    
      '
      ' Case 3
      '
4      lDrawValue = ((ULValues(lIndex) * YMax) / lMaxValue)
      If lDrawValue > 0 Then
        picGraph.Line (X, YMax)-(X, YMax - lDrawValue), vbGreen
        If lMaxPartial < ULValues(lIndex) Then lMaxPartial = ULValues(lIndex)
      End If
      
5      lDrawValue = ((DLValues(lIndex) * YMax) / lMaxValue)
      If lDrawValue > 0 Then
        picGraph.Line (X, YMax)-(X, YMax - lDrawValue), vbYellow
        If lMaxPartial < DLValues(lIndex) Then lMaxPartial = DLValues(lIndex)
      End If
    
    End If
    
    lIndex = (lIndex + 1) Mod UboundDLValues
  Next X
  DoEvents
  If lMaxPartial < lMaxValue Then lMaxValue = lMaxPartial
Exit Sub
ErrH:
  Debug.Print Erl
  Resume Next
End Sub
