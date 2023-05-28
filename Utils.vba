Option Explicit

'''' Global Variables ''''
' GOrigin represents the origin (X,Y,Z = 0,0,0)
Public GOrigin As Point3
' GDrawing represents ThisDrawing.ModelSpace
Public GDrawing As New AcadBlockProxy
' GInvalidValue represents an invalid value
Public Const GInvalidValue As Double = -1
' GInf represents infinity value
Public Const GInf As Double = 655350

Public Const GBlockName_Temp As String = "LZY_BE_AlRIGHT_TEMP"
Public Const GLayerName_Grid As String = "LZY_Grid"
Public Const GLayerName_HalfBreadthPlan As String = "LZY_HalfBreadthPlan"
Public Const GLayerName_SheerPlan As String = "LZY_SheerPlan"
Public Const GLayerName_BodyPlan As String = "LZY_BodyPlan"

' HelloMessage is to avoid this subroutine appearing in the execution panel
Sub InitProgram(HelloMessage As String)
    Debug.Print HelloMessage
    ' Initializes Global Variables
    Set GDrawing.Block = ThisDrawing.ModelSpace
    Set GOrigin = New Point3
    ' This program needs to create some blocks.
    ' Check if the reserved block name exists.
    BlockExists GBlockName_Temp
    ' Create layers
    CreateLayer GLayerName_Grid, acWhite
    CreateLayer GLayerName_HalfBreadthPlan, acCyan
    CreateLayer GLayerName_SheerPlan, acRed
    CreateLayer GLayerName_BodyPlan, acMagenta
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

Sub CreateLayer(LayerName As String, Color As AcColor)
    Dim Layer As AcadLayer
    Set Layer = ThisDrawing.Layers.Add(LayerName)
    Layer.Color = Color
End Sub

Function DrawingArea_Create(BlockName As String) As AcadBlockProxy
    Dim Block As AcadBlock
    Set Block = ThisDrawing.Blocks.Add(GOrigin.ToArray(), BlockName)
    Dim BlockProxy As New AcadBlockProxy
    Set BlockProxy.Block = Block
    Set DrawingArea_Create = BlockProxy
End Function

Sub DrawingArea_DrawAndClean(BlockName As String, InsertPoint As Point3)
    Dim Ref As AcadBlockReference
    Set Ref = ThisDrawing.ModelSpace.InsertBlock(InsertPoint.ToArray(), BlockName, 1, 1, 1, 0)
    Ref.Explode
    Ref.Delete
    Dim Block As AcadBlock
    Set Block = ThisDrawing.Blocks(BlockName)
    Block.Delete
End Sub

