Sub UpdateDocuments()
Application.ScreenUpdating = False
Dim strFolder As String, strFile As String, strDocNm As String
Dim wdDoc As Document, Rng As Range, Sctn As Section, HdFt As HeaderFooter
strDocNm = ActiveDocument.FullName
strFolder = GetFolder
If strFolder = "" Then Exit Sub
strFile = Dir(strFolder & "\*.doc", vbNormal)
While strFile <> ""
  If strFolder & "\" & strFile <> strDocNm Then
    Set wdDoc = Documents.Open(FileName:=strFolder & "\" & strFile, _
      AddToRecentFiles:=False, Visible:=False)
    With wdDoc
      For Each Sctn In .Sections
        For Each HdFt In Sctn.Footers
          If Sctn.Index = 1 Then
            Set Rng = HdFt.Range
            Call ReplaceShape(Rng)
          ElseIf HdFt.LinkToPrevious = False Then
            Set Rng = HdFt.Range
            Call ReplaceShape(Rng)
          End If
        Next
      Next
      '.Close SaveChanges:=True
    End With
  End If
  strFile = Dir()
Wend
Set wdDoc = Nothing
Application.ScreenUpdating = True
End Sub
Sub ReplaceShape(Rng As Range)
Dim sngWdth As Single, SngHght As Single
With Rng
  If .InlineShapes.Count > 0 Then
    Set Rng = .InlineShapes(1).Range
    With .InlineShapes(1)
      sngWdth = .Width
      SngHght = .Height
      .Delete
    End With
    .InlineShapes.AddPicture FileName:="FilePath&Name", LinkToFile:=False, SaveWithDocument:=True, Range:=Rng
    With .InlineShapes(1)
      .Width = sngWdth
      .Height = SngHght
    End With
  End If
End With
End Sub

Function GetFolder() As String
Dim oFolder As Object
GetFolder = ""
Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "Choose a folder", 0)
If (Not oFolder Is Nothing) Then GetFolder = oFolder.Items.Item.Path
Set oFolder = Nothing
End Function