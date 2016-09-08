Attribute VB_Name = "modText"
Public Sub AddChat(rtb As RichTextBox, ParamArray saElements() As Variant)
  Dim arrTmp() As String
  
  With rtb
    .SelStart = Len(.Text)
    .SelLength = 0
    .SelColor = vbWhite
    .SelText = "[" & Time() & "] "
    
    For i = 0 To UBound(saElements) Step 2
      .SelStart = Len(.Text)
      .SelLength = 0
      .SelColor = saElements(i)
      .SelText = saElements(i + 1) & IIf(i + 1 = UBound(saElements), vbNewLine, "")
    Next i
  End With
End Sub

