Sub Main()
    ThisDrawing.PurgeAll ' !!!Warning: FOR DEBUG!!!
    InitProgram
    
    ' TODO: Configurable from file
    Dim FilePath As String
    FilePath = "E:\test\ShipOFF.txt"
    Dim NumericScale As Double
    NumericScale = 1000
    
    Set ShipOFF = ReadDataFromTxtFile(FilePath, NumericScale)
    
    ShipOFF.AddSheerPlane 3000
    ShipOFF.AddSheerPlane 6000
    ShipOFF.AddSheerPlane 9000
    ShipOFF.AddSheerPlane 12000
    ShipOFF.AddSheerPlane 15000
    
    ShipOFF.Breadth = 34000
    ShipOFF.HorizontalPadding = 12000
    
    Dim hbpBlock As AcadBlock
    Set hbpBlock = ThisDrawing.Blocks.Add(GOrigin.ToArray(), GBlockName_HalfBreadthPlan)
    
    Dim hbpBlockProxy As New AcadBlockProxy
    Set hbpBlockProxy.Block = hbpBlock
    
    ShipOFF.DrawHalfBreadthPlanGrid hbpBlockProxy
    ShipOFF.DrawHalfBreadthPlanWaterLine hbpBlockProxy

    Dim pt As New Point3
    pt.XYZ 0, 8000, 0
    Dim Ref As AcadBlockReference
    Set Ref = ThisDrawing.ModelSpace.InsertBlock(pt.ToArray(), GBlockName_HalfBreadthPlan, 1, 1, 1, 0)

End Sub


