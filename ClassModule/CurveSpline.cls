Option Explicit

Public Points As Point3Collection
Public Argument As Variant

Private Sub Class_Initialize()
    Set Points = New Point3Collection
End Sub

Public Sub Add(ByVal point As Point3)
    Points.Add point
End Sub

Public Sub AddXYZ(ByVal x As Double, ByVal y As Double, ByVal z As Double)
    Points.AddXYZ x, y, z
End Sub

Public Sub Reverse()
    Points.Reverse
End Sub
