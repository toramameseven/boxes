Private Sub PrintIDs()
 
  Dim CBTN As CommandBarButton
  Dim CBR As CommandBar
  On Error Resume Next

  For Each CBR In Application.CommandBars
      For Each CBTN In CBR.Controls
          Selection.TypeText CBR.Name & ": " & CBTN.ID & " - " & CBTN.Caption
          Selection.TypeParagraph
      Next
  Next
 
End Sub
