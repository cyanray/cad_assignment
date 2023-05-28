Option Explicit

' ElementType: CurveSpline
Private m_WaterLines As Collection
' ElementType: CurveSpline
Private m_SheerLines As Collection
' ElementType: CurveSpline
Private m_BodyLines As Collection

' ElementType: Station
Private m_Stations As Collection
' ElementType: Double
Private m_WaterPlane As Collection
' ElementType: Double
Private m_SheerPlane As Collection

Private m_Breadth As Double

Private m_Depth As Double

Private m_HorizontalPadding As Double

Private m_TextHeight As Double

Private Sub Class_Initialize()
    Set m_WaterLines = New Collection
    Set m_SheerLines = New Collection
    Set m_BodyLines = New Collection
    Set m_Stations = New Collection
    Set m_WaterPlane = New Collection
    Set m_SheerPlane = New Collection
    m_Breadth = GInvalidValue
    m_Depth = GInvalidValue
    HorizontalPadding = 0
End Sub

Private Property Get HorizontalLineLeft() As double
    HorizontalLineLeft = m_Stations(1).StationOffset - m_HorizontalPadding
End Property

Private Property Get HorizontalLineRight() As double
    HorizontalLineRight = m_Stations(m_Stations.Count).StationOffset + m_HorizontalPadding
End Property

Public Property Get HalfBreadth() As Double
    HalfBreadth = m_Breadth / 2
End Property

Public Property Get Breadth() As Double
    Breadth = m_Breadth
End Property

Public Property Let Breadth(ByVal Value As Double)
    m_Breadth = Value
End Property

Public Property Get Depth() As Double
    Depth = m_Depth
End Property

Public Property Let Depth(ByVal Value As Double)
    m_Depth = Value
End Property

Public Property Get HorizontalPadding() As Double
    HorizontalPadding = m_HorizontalPadding
End Property

Public Property Let HorizontalPadding(ByVal Value As Double)
    m_HorizontalPadding = Value
End Property

Public Property Get TextHeight() As Double
    TextHeight = m_TextHeight
End Property

Public Property Let TextHeight(ByVal Value As Double)
    m_TextHeight = Value
End Property

Public Property Get WaterLines() As Collection
    Set WaterLines = m_WaterLines
End Property

Public Property Get SheerLines() As Collection
    Set SheerLines = m_SheerLines
End Property

Public Property Get BodyLines() As Collection
    Set BodyLines = m_BodyLines
End Property

Public Property Get SheerPlanes() As Collection
    Set SheerPlanes = m_SheerPlane
End Property

Public Property Get Stations() As Collection
    Dim result As Collection
    Set result = New Collection
    Dim i As Integer
    For i = 1 To m_Stations.Count
        result.Add m_Stations(i).StationOffset
    Next i
    Set Stations = result
End Property

Public Sub AddWaterLine(wl As CurveSpline)
    m_WaterLines.Add wl
End Sub

Public Sub AddSheerLine(sl As CurveSpline)
    m_SheerLines.Add sl
End Sub

Public Sub AddBodyLine(bl As CurveSpline)
    m_BodyLines.Add bl
End Sub

Public Sub AddStation(sta As Station)
    ' TODO: Ensure the elements are sorted.
    m_Stations.Add sta
End Sub

Public Sub AddStationByValue(staNumber As Double, staOffset As Double)
    ' TODO: Ensure the elements are sorted.
    Dim sta As New Station
    sta.StationNumber = staNumber
    sta.StationOffset = staOffset
    AddStation sta
End Sub

Public Sub AddWaterPlane(sheerPlane As Double)
    ' TODO: Ensure the elements are sorted.
    m_WaterPlane.Add sheerPlane
End Sub

Public Sub AddSheerPlane(sheerPlane As Double)
    ' TODO: Ensure the elements are sorted.
    m_SheerPlane.Add sheerPlane
End Sub

Public Sub DrawHalfBreadthPlanWaterLine(blockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    Dim i As Integer
    For i = 1 To m_WaterLines.Count
        blockProxy.AddSpline m_WaterLines(i).Points, GOrigin, GOrigin, LayerName
    Next i
End Sub

Public Sub DrawHalfBreadthPlanGrid(blockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    ' Throw exception if m_Breadth invalid
    If m_Breadth = GInvalidValue Then
        Err.Raise 10001, "ShipOffsets", "m_Breadth is invalid."
    End If

    Dim leftOffset As Double
    Dim rightOffset As Double
    leftOffset = m_Stations(1).StationOffset
    rightOffset = m_Stations(m_Stations.Count).StationOffset
    Dim startLength As Double: startLength = leftOffset - m_HorizontalPadding
    Dim endLength As Double: endLength = rightOffset + m_HorizontalPadding

    Dim i As Integer
    For i = 1 To m_Stations.Count
        Dim sta As Station: Set sta = m_Stations(i)
        blockProxy.AddLineXYZXYZ sta.StationOffset, 0, 0, sta.StationOffset, HalfBreadth, 0, LayerName
    Next i
    For i = 1 To m_SheerPlane.Count
        Dim sheerPlane As Double: sheerPlane = m_SheerPlane(i)
        If sheerPlane = 0 Or sheerPlane = HalfBreadth Then
            GoTo ContinueLoop
        End If
        blockProxy.AddLineXYZXYZ startLength, sheerPlane, 0, endLength, sheerPlane, 0, LayerName
ContinueLoop:
    Next i
    ' BaseLine
    blockProxy.AddLineXYZXYZ startLength, 0, 0, endLength, 0, 0, LayerName
    ' TopLine
    blockProxy.AddLineXYZXYZ startLength, HalfBreadth, 0, endLength, HalfBreadth, 0, LayerName
End Sub

Public Sub DrawSheerPlanSheerLine(blockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    Dim i As Integer
    For i = 1 To m_SheerLines.Count
        Dim ln As AcadSpline
        ' Patch for parallel middle body
        Dim slOrigin As CurveSpline: Set slOrigin = m_SheerLines(i)
        ' Find the index of point that Y = 0
        Dim Idx As Integer
        For Idx = 1 To slOrigin.Points.Count
            Dim pt As Point3: Set pt = slOrigin.Points.Item(Idx)
            If pt.Y = 0 Then
                Exit For
            End If
        Next Idx
        ' If found
        If Idx < slOrigin.Points.Count Then
            Dim x1 As Double: x1 = slOrigin.Points.Item(Idx).X
            Dim x2 As Double: x2 = slOrigin.Points.Item(Idx + 1).X
            Dim step As Double: step = (x2 - x1) / 16
            Dim sl As CurveSpline: Set sl = New CurveSpline
            Dim j As Integer
            For j = 1 To Idx
                sl.Add slOrigin.Points.Item(j)
            Next j
            For j = 1 To 16
                sl.AddXYZ x1 + j * step, 0, 0
            Next j
            For j = Idx + 1 To slOrigin.Points.Count
                sl.Add slOrigin.Points.Item(j)
            Next j
            Set ln = blockProxy.AddSpline(sl.Points, GOrigin, GOrigin)
        Else
            Set ln = blockProxy.AddSpline(slOrigin.Points, GOrigin, GOrigin)
        End If
        ln.Layer = LayerName
    Next i
End Sub

Public Sub DrawSheerPlanGrid(blockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    ' Throw exception if m_Depth invalid
    If m_Depth = GInvalidValue Then
        Err.Raise 10001, "ShipOffsets", "m_Depth is invalid."
    End If

    Dim leftOffset As Double
    Dim rightOffset As Double
    leftOffset = m_Stations(1).StationOffset
    rightOffset = m_Stations(m_Stations.Count).StationOffset
    Dim startLength As Double: startLength = leftOffset - m_HorizontalPadding
    Dim endLength As Double: endLength = rightOffset + m_HorizontalPadding

    Dim i As Integer
    For i = 1 To m_Stations.Count
        Dim sta As Station: Set sta = m_Stations(i)
        blockProxy.AddLineXYZXYZ sta.StationOffset, 0, 0, sta.StationOffset, m_Depth, 0, LayerName
    Next i
    For i = 1 To m_WaterPlane.Count
        Dim waterPlane As Double: waterPlane = m_WaterPlane(i)
        If waterPlane = 0 Or waterPlane = m_Depth Then
            GoTo ContinueLoop
        End If
        blockProxy.AddLineXYZXYZ startLength, waterPlane, 0, endLength, waterPlane, 0, LayerName
ContinueLoop:
    Next i
    ' BaseLine
    blockProxy.AddLineXYZXYZ startLength, 0, 0, endLength, 0, 0, LayerName
    ' TopLine
    blockProxy.AddLineXYZXYZ startLength, m_Depth, 0, endLength, m_Depth, 0, LayerName
End Sub

Public Sub DrawBodyPlanBodyLine(blockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    Dim i As Integer
    For i = 1 To m_BodyLines.Count
        blockProxy.AddSpline m_BodyLines(i).Points, GOrigin, GOrigin, LayerName
    Next i
End Sub

Public Sub DrawBodyPlanGrid(blockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    ' Throw exception if m_Depth or m_Breadth invalid
    If m_Depth = GInvalidValue Then
        Err.Raise 10001, "ShipOffsets", "m_Depth is invalid."
    End If
    If m_Breadth = GInvalidValue Then
        Err.Raise 10001, "ShipOffsets", "m_Breadth is invalid."
    End If

    Dim leftOffset As Double
    Dim rightOffset As Double
    leftOffset = -HalfBreadth - m_HorizontalPadding
    rightOffset = HalfBreadth + m_HorizontalPadding

    ' draw water plane and sheer plane
    Dim i As Integer
    For i = 1 To m_WaterPlane.Count
        Dim waterPlane As Double: waterPlane = m_WaterPlane(i)
        blockProxy.AddLineXYZXYZ leftOffset, waterPlane, 0, rightOffset, waterPlane, 0, LayerName
    Next i
    For i = 1 To m_SheerPlane.Count
        Dim sheerPlane As Double: sheerPlane = m_SheerPlane(i)
        blockProxy.AddLineXYZXYZ sheerPlane, 0, 0, sheerPlane, Depth, 0, LayerName
        blockProxy.AddLineXYZXYZ -sheerPlane, 0, 0, -sheerPlane, Depth, 0, LayerName
    Next i
    ' BaseLine
    blockProxy.AddLineXYZXYZ leftOffset, 0, 0, rightOffset, 0, 0, LayerName
    blockProxy.AddLineXYZXYZ 0,0,0, 0, Depth, 0, LayerName
    ' blockProxy.AddLineXYZXYZ -HalfBreadth, Depth, 0, HalfBreadth, Depth, 0, LayerName
    blockProxy.AddLineXYZXYZ -HalfBreadth, 0, 0, -HalfBreadth, Depth, 0, LayerName
    blockProxy.AddLineXYZXYZ HalfBreadth, 0, 0, HalfBreadth, Depth, 0, LayerName
End Sub

Public Sub DrawHalfBreadthPlanText(BlockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    Dim pt As Point3
    Dim i As Integer
    For i = 1 To m_Stations.Count
        Dim sta As Station: Set sta = m_Stations(i)
        Set pt = New Point3
        pt.XYZ sta.StationOffset, 0, 0
        BlockProxy.AddText CStr(Abs(sta.StationNumber)), pt, TextHeight, acAlignmentTopCenter, LayerName
    Next i
    For i = 1 To m_SheerPlane.Count
        Dim sheerPlane As Double: sheerPlane = m_SheerPlane(i)
        ' skip sheer plane at 0
        If sheerPlane = 0 Then
            GoTo ContinueLoop
        End If
        Set pt = New Point3
        pt.XYZ HorizontalLineLeft, sheerPlane, 0
        BlockProxy.AddText CStr(Abs(sheerPlane))&"BL", pt, TextHeight, acAlignmentLeft, LayerName
        pt.XYZ HorizontalLineRight, sheerPlane, 0
        BlockProxy.AddText CStr(Abs(sheerPlane))&"BL", pt, TextHeight, acAlignmentRight, LayerName
ContinueLoop:
    Next i
    'Base line
    pt.XYZ HorizontalLineLeft, 0, 0
    BlockProxy.AddText "BL", pt, TextHeight, acAlignmentLeft, LayerName
    pt.XYZ HorizontalLineRight, 0, 0
    BlockProxy.AddText "BL", pt, TextHeight, acAlignmentRight, LayerName
End Sub

Public Sub DrawSheerPlanText(BlockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    Dim pt As Point3
    Dim i As Integer
    For i = 1 To m_Stations.Count
        Dim sta As Station: Set sta = m_Stations(i)
        Set pt = New Point3
        pt.XYZ sta.StationOffset, 0, 0
        BlockProxy.AddText CStr(Abs(sta.StationNumber)), pt, TextHeight, acAlignmentTopCenter, LayerName
    Next i
    For i = 1 To m_WaterPlane.Count
        Dim waterPlane As Double: waterPlane = m_WaterPlane(i)
        ' skip water plane = 0
        If waterPlane = 0 Then
            GoTo ContinueLoop
        End If
        Set pt = New Point3
        pt.XYZ HorizontalLineLeft, waterPlane, 0
        BlockProxy.AddText CStr(Abs(waterPlane))&"WL", pt, TextHeight, acAlignmentLeft, LayerName
        pt.XYZ HorizontalLineRight, waterPlane, 0
        BlockProxy.AddText CStr(Abs(waterPlane))&"WL", pt, TextHeight, acAlignmentRight, LayerName
ContinueLoop:
    Next i
    'Base line
    pt.XYZ HorizontalLineLeft, 0, 0
    BlockProxy.AddText "BL", pt, TextHeight, acAlignmentLeft, LayerName
    pt.XYZ HorizontalLineRight, 0, 0
    BlockProxy.AddText "BL", pt, TextHeight, acAlignmentRight, LayerName
End Sub

Public Sub DrawBodyPlanText(BlockProxy As AcadBlockProxy, Optional LayerName As String = "0")
    Dim pt As Point3
    Dim i As Integer

    Dim leftOffset As Double: leftOffset = -HalfBreadth - m_HorizontalPadding
    Dim rightOffset As Double: rightOffset = HalfBreadth + m_HorizontalPadding

    For i = 1 To m_SheerPlane.Count
        Dim sheerPlane As Double: sheerPlane = m_SheerPlane(i)
        ' skip sheer plane at 0
        If sheerPlane = 0 Then
            GoTo ContinueLoop1
        End If
        Set pt = New Point3
        pt.XYZ sheerPlane, 0, 0
        BlockProxy.AddText CStr(Abs(sheerPlane)), pt, TextHeight, acAlignmentTopCenter, LayerName
        pt.XYZ -sheerPlane, 0, 0
        BlockProxy.AddText CStr(Abs(sheerPlane)), pt, TextHeight, acAlignmentTopCenter, LayerName
ContinueLoop1:
    Next i

    For i = 1 To m_WaterPlane.Count
        Dim waterPlane As Double: waterPlane = m_WaterPlane(i)
        ' skip water plane = 0
        If waterPlane = 0 Then
            GoTo ContinueLoop2
        End If
        Set pt = New Point3
        pt.XYZ leftOffset, waterPlane, 0
        BlockProxy.AddText CStr(Abs(waterPlane))&"WL", pt, TextHeight, acAlignmentLeft, LayerName
        pt.XYZ rightOffset, waterPlane, 0
        BlockProxy.AddText CStr(Abs(waterPlane))&"WL", pt, TextHeight, acAlignmentRight, LayerName
ContinueLoop2:
    Next i
    ' Base line
    pt.XYZ 0,0,0
    BlockProxy.AddText "CL", pt, TextHeight, acAlignmentTopCenter, LayerName
    pt.XYZ leftOffset, 0, 0
    BlockProxy.AddText "BL", pt, TextHeight, acAlignmentLeft, LayerName
    pt.XYZ rightOffset, 0, 0
    BlockProxy.AddText "BL", pt, TextHeight, acAlignmentRight, LayerName
End Sub