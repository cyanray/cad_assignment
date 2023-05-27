Option Explicit

Function ReadDataFromTxtFile(FilePath As String, Optional NumericScale As Double = 1#) As ShipOffsets
    Dim Result As New ShipOffsets
    
    ' TODO: Determine from file
    Const MaxRow As Integer = 14
    Const MaxCol As Integer = 29

    Dim matrix(1 To MaxRow, 1 To MaxCol) As Double
    
    ' Read matrix from file
    Open FilePath For Input As #1
    Dim i As Integer, j As Integer
    For i = 1 To MaxRow
        Dim line As String
        Line Input #1, line
        Dim numbers() As String
        numbers = Split(line, " ")
        For j = 1 To MaxCol
            matrix(i, j) = CDbl(numbers(j - 1))
            If Not i = 1 And Not matrix(i, j) = GInvalidValue Then
                matrix(i, j) = matrix(i, j) * NumericScale
            End If
        Next j
    Next i
    Close #1

    ' Print matrix
    For i = 1 To MaxRow
        For j = 1 To MaxCol
            Debug.Print matrix(i, j);
        Next j
        Debug.Print
    Next i

    ' The first and last line are station data
    For i = 1 To MaxCol
        If matrix(1, i) = GInvalidValue Or matrix(MaxRow, i) = GInvalidValue Then
            GoTo ContinueLoop
        End If

        Result.AddStationByValue matrix(1, i), matrix(MaxRow, i)
ContinueLoop:
    Next i

    ' The first column are water planes
    For i = 2 To MaxRow
        If matrix(i, 1) = GInvalidValue Then
            GoTo ContinueLoop1
        End If
        Result.AddWaterPlane matrix(i, 1)
ContinueLoop1:
    Next i

    ' Parsing waterlines
    For i = 2 To MaxRow - 2
        Dim wl As CurveSpline
        Set wl = New CurveSpline
        wl.Argument = matrix(i, 1)
        If Not matrix(i, 2) = GInvalidValue Then
            wl.AddXYZ matrix(i, 2), 0, 0
        End If

        For j = 3 To MaxCol - 1
            If matrix(i, j) = GInvalidValue Then
                GoTo ContinueLoop2
            End If
            wl.AddXYZ matrix(MaxRow, j), matrix(i, j), 0
ContinueLoop2:
        Next j

        If Not matrix(i, MaxCol) = GInvalidValue Then
            wl.AddXYZ matrix(i, MaxCol), 0, 0
        End If
        Result.AddWaterLine wl
    Next i

    ' Work done.
    Set ReadDataFromTxtFile = Result
End Function

Sub GenerateSheerLinesFromWaterLines(ByRef offset As ShipOffsets)
    ReDim Result(1 To offset.SheerPlanes.Count) As CurveSpline
    Dim i As Integer
    For i = LBound(Result) To UBound(Result)
        Set Result(i) = New CurveSpline
    Next i

    IntersectWithWaterLines offset, Result, False

    Dim Item
    For Each Item In Result
        Item.Reverse
    Next Item

    IntersectWithWaterLines offset, Result, True

    For i = LBound(Result) To UBound(Result)
        offset.AddSheerLine Result(i)
    Next i

End Sub

Sub IntersectWithWaterLines(ByRef offset As ShipOffsets, ByRef Result() As CurveSpline, Flag As Boolean)
    Dim IdxWaterLine As Integer
    Dim IdxSheerPlane As Integer
    Dim tempBlock As AcadBlock
    Set tempBlock = ThisDrawing.Blocks.Add(GOrigin.ToArray(), GBlockName_Temp)
    Dim tempProxy As New AcadBlockProxy: Set tempProxy.Block = tempBlock
        For IdxWaterLine = 1 To offset.WaterLines.Count
        Dim wl As AcadSpline
        Set wl = tempProxy.AddSpline(offset.WaterLines(IdxWaterLine).Points, GOrigin, GOrigin)
        For IdxSheerPlane = 1 To offset.SheerPlanes.Count
            Dim sp As AcadLine
            If Flag Then
                Set sp = tempProxy.AddLineXYZXYZ(0, offset.SheerPlanes(IdxSheerPlane), 0, GInf, offset.SheerPlanes(IdxSheerPlane), 0)
            Else
                Set sp = tempProxy.AddLineXYZXYZ(-GInf, offset.SheerPlanes(IdxSheerPlane), 0, 0, offset.SheerPlanes(IdxSheerPlane), 0)
            End If
            Dim pointsArray
            pointsArray = sp.IntersectWith(wl, acExtendNone)
            Dim Points As Point3Collection: Set Points = New Point3Collection
            Points.AddFromArray pointsArray
            If Points.Count = 1 And offset.WaterLines(IdxWaterLine).Argument <> GInvalidValue Then
                Result(IdxSheerPlane).AddXYZ Points.Item(1).x, offset.WaterLines(IdxWaterLine).Argument, 0
            End If
            sp.Delete
        Next IdxSheerPlane
        wl.Delete
    Next IdxWaterLine
    tempBlock.Delete
End Sub