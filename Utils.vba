Option Explicit

'''' Global Variables ''''
' GOrigin represents the origin (X,Y,Z = 0,0,0)
Public GOrigin As Point3
' GDrawing represents ThisDrawing.ModelSpace
Public GDrawing As New AcadBlockProxy
' GInvalidValue represents an invalid value
Public Const GInvalidValue As Double = -1
' GInf represents infinity value
Public Const GInf As Integer = 4096

Public Const GBlockName_Temp As String = "LZY_TEMP"
Public Const GBlockName_HalfBreadthPlan As String = "LZY_HalfBreadthPlan"
Public Const GBlockName_SheerPlan As String = "LZY_SheerPlan"
Public Const GBlockName_BodyPlan As String = "LZY_BodyPlan"


Sub InitProgram()
    ' Initializes Global Variables
    Set GDrawing.Block = ThisDrawing.ModelSpace
    Set GOrigin = New Point3
    ' This program needs to create some blocks.
    ' Check if the reserved block name exists.
    BlockExists GBlockName_Temp
    BlockExists GBlockName_HalfBreadthPlan
    BlockExists GBlockName_SheerPlan
    BlockExists GBlockName_BodyPlan
End Sub


Sub BlockExists(BlockName As String)
    Dim Exists As Boolean
    Exists = False
    On Error Resume Next
        Dim Blk As AcadBlock
        Set Blk = ThisDrawing.Blocks(BlockName)
        Exists = Not Blk Is Nothing
    On Error GoTo 0
    If Exists Then
        Err.Raise 20001, "Utils", "Block '" & BlockName & "' is exists."
    End If
End Sub

Function Max(a As Variant, b As Variant) As Variant
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

Function Min(a As Variant, b As Variant) As Variant
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function