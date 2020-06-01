' =================================================================================================================
' Open Doc - DONE
Sub OpenDoc()
' Dim appWord As Object
' Set appWord = CreateObject("Word.Application")
    Documents.Open FileName:="C:\Users\abhijit\OneDrive\Desktop\fiverr_june\notities-dynamica-a-orig.docx", ReadOnly:=True
End Sub
' =================================================================================================================
' Count the total no. of pages
Sub get_tot_pgs_doc()
    tot_pgs = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
    MsgBox (tot_pgs)
End Sub
' =================================================================================================================
' iterate over odd pages starting from 3, 5, 7, ....
' replace `.ComputeStatistics(wdStatisticPages)` with `tot_pages`
Sub loop_pages()
    Dim i As Long, j As Long, Rng As Range
    With ActiveDocument
      For i = 3 To .ComputeStatistics(wdStatisticPages) Step 2
        Set Rng = ActiveDocument.GoTo(What:=wdGoToPage, Name:=i)
        Set Rng = Rng.GoTo(What:=wdGoToBookmark, Name:="\page")
        With Rng
          For j = 1 To .ShapeRange.Count
            If .ShapeRange(i).Name = "Barcode" Then
              MsgBox "Found on page: " & i
            End If
          Next
        End With
      Next
    End With
End Sub 

' =================================================================================================================
' replace image - DONE
' Here, it changes the 1st image, as given `InlineShapes(1)`
Sub replaceImage()

    Dim originalImage As InlineShape
    Dim newImage As InlineShape

    Set originalImage = ActiveDocument.InlineShapes(1)

    Dim imageControl As ContentControl

    If originalImage.Range.ParentContentControl Is Nothing Then
        Set imageControl = ActiveDocument.ContentControls.Add(wdContentControlPicture, originalImage.Range)
    Else
        Set imageControl = originalImage.Range.ParentContentControl
    End If

    Dim imageW As Long
    Dim imageH As Long
    imageW = originalImage.Width
    imageH = originalImage.Height

    originalImage.Delete

    Dim imagePath As String
    imagePath = "C:\Users\abhijit\OneDrive\Desktop\fiverr_june\white.png"
    ActiveDocument.InlineShapes.AddPicture imagePath, False, True, imageControl.Range

    With imageControl.Range.InlineShapes(1)
        .Height = imageH
        .Width = imageW
    End With

End Sub
' =================================================================================================================

