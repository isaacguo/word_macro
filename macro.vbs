Sub AllPictSize()
        Dim targetHeight As Integer
    Dim oShp As Shape
    Dim oILShp As InlineShape
 
    targetWidth = 16
 
    For Each oShp In ActiveDocument.Shapes
        If oShp.Width > CentimetersToPoints(targetWidth) Then
           With oShp
                .Height = AspectHt(.Height, .Width, CentimetersToPoints(targetWidth))
                .Width = CentimetersToPoints(targetWidth)
            End With

        End If
        
    Next
 
    For Each oILShp In ActiveDocument.InlineShapes
        If oILShp.Width > CentimetersToPoints(targetWidth) Then
            With oILShp
                .Height = AspectHt(.Height, .Width, CentimetersToPoints(targetWidth))
                .Width = CentimetersToPoints(targetWidth)
            End With
        End If
        
    Next

 
 
 
End Sub
 

 
Private Function AspectHt(ByVal origHt As Long, ByVal origWd As Long, ByVal newWd As Long) As Long
    If origWd <> 0 Then
        AspectHt = (CSng(origHt) / CSng(origWd)) * newWd
    Else
        AspectHt = 0
    End If
End Function

Sub AutoFitWindowForAllTables()
  If ActiveDocument.Tables.Count > 0 Then
    Dim objTable As Object

    Application.Browser.Target = wdBrowseTable
    For Each objTable In ActiveDocument.Tables
      objTable.AutoFitBehavior (wdAutoFitWindow)
    Next
  End If
End Sub



