Function ConvertPointToCm(ByVal pnt As Double) As Double
    ConvertPointToCm = pnt * 0.03527778
End Function

Function ConvertCmToPoint(ByVal cm As Double) As Double
    ConvertCmToPoint = cm * 28.34646
End Function

Sub SplitObject()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    numRows = CInt(InputBox("Number of rows", "Rows", 1))
    numCols = CInt(InputBox("Number of columns", "Columns", 1))

    spacingRows = ConvertCmToPoint(CDbl(InputBox("Spacing rows (cm)", "Spacing Rows", 0)))
    spacingCols = ConvertCmToPoint(CDbl(InputBox("Spacing cols (cm)", "Spacing Columns", 0)))

    If numRows <= 0 Or numCols <= 0 Then
        Debug.Print "Invalid number of splits"
        Exit Sub
    End If
    
    If MyDocument.Selection.HasChildShapeRange Then
        If MyDocument.Selection.ChildShapeRange.Count > 1 Then
            MsgBox "Select only one object to split."
            Exit Sub
        Else
            Set obj = MyDocument.Selection.ChildShapeRange(1)
        End If
    Else
        If MyDocument.Selection.ShapeRange.Count > 1 Then
            MsgBox "Select only one object to split."
            Exit Sub
        Else
            Set obj = MyDocument.Selection.ShapeRange(1)
        End If
    End If

    objWidth = obj.Width
    objHeight = obj.Height
    
    If objWidth = 0 Or objHeight = 0 Then
        Debug.Print "Invalid object dimensions."
        Exit Sub
    End If
    
    splitWidth = objWidth / numCols
    splitHeight = objHeight / numRows
    
    ReDim splitArray(1 To numRows, 1 To numCols)
    
    k = 1
    For i = 1 To numRows
        l = 1
        For j = 1 To numCols
            Set splitArray(i, j) = obj.Duplicate
            splitArray(i, j).Top = obj.Top + ((i - 1) * (splitHeight + spacingRows))
            splitArray(i, j).Left = obj.Left + ((j - 1) * (splitWidth + spacingCols))
            splitArray(i, j).Width = splitWidth
            splitArray(i, j).Height = splitHeight
            l = l + 1
        Next j
        k = k + 1
    Next i

    obj.Delete
    
End Sub