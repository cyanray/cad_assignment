Sub Main()
    ' ThisDrawing.PurgeAll ' !!!Warning: FOR DEBUG!!!
    InitProgram "Program Start"
    
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
    Call GenerateBodyLinesFromWaterLines(ShipOff, ShipOff.Stations)
    
    Dim pt As Point3: Set pt = New Point3
    Dim proxy As AcadBlockProxy
    
    Set proxy = DrawingArea_Create(GBlockName_Temp)
        pt.XYZ 0, 0, 0
        ShipOff.DrawHalfBreadthPlanGrid proxy, GLayerName_Grid
        ShipOff.DrawHalfBreadthPlanWaterLine proxy, GLayerName_HalfBreadthPlan
    Call DrawingArea_DrawAndClean(GBlockName_Temp, pt)

    Set proxy = DrawingArea_Create(GBlockName_Temp)
        pt.XYZ 0, 30000, 0
        ShipOff.DrawSheerPlanGrid proxy, GLayerName_Grid
        ShipOff.DrawSheerPlanSheerLine proxy, GLayerName_SheerPlan
    Call DrawingArea_DrawAndClean(GBlockName_Temp, pt)

    Set proxy = DrawingArea_Create(GBlockName_Temp)
        pt.XYZ 0, 60000, 0
        ShipOff.DrawBodyPlanGrid proxy, GLayerName_Grid
        ShipOff.DrawBodyPlanBodyLine proxy, GLayerName_BodyPlan
    Call DrawingArea_DrawAndClean(GBlockName_Temp, pt)

End Sub

