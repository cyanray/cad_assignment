Sub Main()
    ' ThisDrawing.PurgeAll ' !!!Warning: FOR DEBUG!!!
    InitProgram
    
    ' TODO: Configurable from file
    Dim FilePath As String
    FilePath = "E:\test\ShipOFF.txt"
    Dim NumericScale As Double
    NumericScale = 1000
    
    Dim ShipOff As ShipOffsets
    Set ShipOff = ReadDataFromTxtFile(FilePath, NumericScale)
    
    ShipOff.AddSheerPlane 3000
    ShipOff.AddSheerPlane 6000
    ShipOff.AddSheerPlane 9000
    ShipOff.AddSheerPlane 12000
    ShipOff.AddSheerPlane 15000
    
    ShipOff.Breadth = 34000
    ShipOff.Depth = 19000
    ShipOff.HorizontalPadding = 12000
    
    Call GenerateSheerLinesFromWaterLines(ShipOff)
    
    Dim pt As Point3
    Set pt = New Point3

    Dim hbpBlock As AcadBlock
    Set hbpBlock = ThisDrawing.Blocks.Add(GOrigin.ToArray(), GBlockName_HalfBreadthPlan)
    Dim hbpBlockProxy As New AcadBlockProxy
    Set hbpBlockProxy.Block = hbpBlock
    
    ShipOff.DrawHalfBreadthPlanGrid hbpBlockProxy
    ShipOff.DrawHalfBreadthPlanWaterLine hbpBlockProxy

    Dim hbpRef As AcadBlockReference
    Set hbpRef = ThisDrawing.ModelSpace.InsertBlock(GOrigin.ToArray(), GBlockName_HalfBreadthPlan, 1, 1, 1, 0)
    hbpRef.Explode
    hbpRef.Delete
    hbpBlock.Delete

    Dim spBlock As AcadBlock
    Set spBlock = ThisDrawing.Blocks.Add(GOrigin.ToArray(), GBlockName_SheerPlan)
    Dim spBlockProxy As New AcadBlockProxy
    Set spBlockProxy.Block = spBlock

    ShipOff.DrawSheerPlanGrid spBlockProxy
    ShipOff.DrawSheerPlanSheerLine spBlockProxy

    pt.XYZ 0, 30000, 0
    Dim spRef As AcadBlockReference
    Set spRef = ThisDrawing.ModelSpace.InsertBlock(pt.ToArray(), GBlockName_SheerPlan, 1, 1, 1, 0)
    spRef.Explode
    spRef.Delete
    spBlock.Delete
    
End Sub

