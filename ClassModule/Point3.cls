Option Explicit

Public X As Double, Y As Double, Z As Double

Public Sub XYZ(xval As Double, yval As Double, zval As Double)
X = xval: Y = yval: Z = zval
End Sub

Public Function ToArray() As Double()
    Dim Result(2) As Double
    Result(0) = X: Result(1) = Y: Result(2) = Z
    ToArray = Result
End Function

