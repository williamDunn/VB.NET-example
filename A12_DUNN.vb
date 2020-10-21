Sub A12_DUNN()

Dim names As Collection
Set names = learnNames

Selection.Font.Size = 15
Selection.Font.Name = "Arial Black"
Selection.Font.Color = wdColorBlue

Dim i As Integer
    For i = 1 To names.Count
    Selection.TypeText (names(i))
    Selection.TypeParagraph
    Next
    
Dim counted As String
counted = CStr(names.Count)
    
With ActiveDocument.Sections(1)
    .Footers(wdHeaderFooterPrimary).Range.Text = counted & " names."
End With

End Sub


Function learnNames() As Collection
    Dim names As New Collection
    names.Add ("Will")
    names.Add ("Mom")
    names.Add ("Sister")
    names.Add ("Oreo")
    
    Set learnNames = names
End Function
