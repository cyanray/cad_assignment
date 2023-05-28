Option Explicit

Private m_Collection As Collection

Private Sub Class_Initialize()
    Set m_Collection = New Collection
End Sub

Public Sub Add(point As Point3)
    m_Collection.Add point
End Sub

Public Sub AddXYZ(x As Double, y As Double, z As Double)
    Dim point As New Point3
    point.XYZ x, y, z
    Add point
End Sub

Public Sub AddFromArray(arr As Variant)
    Dim Idx As Integer
    Dim point As Point3
    For Idx = 0 To UBound(arr) Step 3
        Call AddXYZ((arr(Idx + 0)), (arr(Idx + 1)), (arr(Idx + 2)))
    Next
End Sub

Public Function Count() As Integer
    Count = m_Collection.Count
End Function

Public Function Item(index As Integer) As Point3
    Set Item = m_Collection.Item(index)
End Function

Public Function ToArray() As Double()
    Dim UpperBnd As Integer
    UpperBnd = 3 * m_Collection.Count - 1
    ReDim Result(UpperBnd) As Double
    Dim Idx As Integer
    Dim point As Point3
    For Idx = 1 To m_Collection.Count
        Set point = m_Collection.Item(Idx)
        Result((Idx - 1) * 3 + 0) = point.x
        Result((Idx - 1) * 3 + 1) = point.y
        Result((Idx - 1) * 3 + 2) = point.z
    Next
    ToArray = Result
End Function

Public Sub Reverse()
    Dim NewCol As Collection
    Set NewCol = New Collection
    Dim obj
    For Each obj In m_Collection
        If NewCol.Count > 0 Then
            NewCol.Add Item:=obj, before:=1
        Else
            NewCol.Add Item:=obj
        End If
    Next
    Set m_Collection = NewCol
End Sub
