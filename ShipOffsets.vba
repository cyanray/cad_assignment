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

Private m_HorizontalPadding As Double

Private Sub Class_Initialize()
    Set m_WaterLines = New Collection
    Set m_SheerLines = New Collection
    Set m_BodyLines = New Collection
    Set m_Stations = New Collection
    Set m_WaterPlane = New Collection
    Set m_SheerPlane = New Collection
    m_Breadth = GInvalidValue
    HorizontalPadding = 0
End Sub

Public Property Get HalfBreadth() As Double
    HalfBreadth = m_Breadth / 2
End Property

Public Property Get Breadth() As Double
    Breadth = m_Breadth
End Property

Public Property Let Breadth(ByVal Value As Double)
    m_Breadth = Value
End Property

Public Property Get HorizontalPadding() As Double
    HorizontalPadding = m_HorizontalPadding
End Property

Public Property Let HorizontalPadding(ByVal Value As Double)
    m_HorizontalPadding = Value
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


Public Sub AddWaterLine(wl As CurveSpline)
    m_WaterLines.Add wl
End Sub

Public Sub AddSheerLine(sl As CurveSpline)
    m_SheerLines.Add sl
End Sub

Public Sub AddStation(sta As Station)
    ' TODO: Ensure the elements are sorted.
    m_Stations.Add sta
End Sub

Public Sub AddStationByValue(staNumber As double, staOffset As Double)
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

Public Sub DrawHalfBreadthPlanWaterLine(blockProxy As AcadBlockProxy)
    Dim i As Integer
    For i = 1 To m_WaterLines.Count
        Dim wl As CurveSpline : Set wl = m_WaterLines(i)
        Dim ln As AcadSpline
        Set ln = blockProxy.AddSpline(wl.Points, GOrigin, GOrigin)
        ln.color = acCyan
    Next i
End Sub

Public Sub DrawHalfBreadthPlanGrid(blockProxy As AcadBlockProxy)
    ' Throw exception if m_Breadth or m_LeftLength invalid
    If m_Breadth = GInvalidValue Then
        Err.Raise 10001, "ShipOffsets", "m_Breadth is invalid."
    End If

    Dim staLeftOffset As Double
    Dim staRightOffset As Double
    staLeftOffset = m_Stations(1).StationOffset
    staRightOffset = m_Stations(m_Stations.Count).StationOffset
    Dim startLength As Double : startLength = staLeftOffset - m_HorizontalPadding
    Dim endLength As Double : endLength = staRightOffset + m_HorizontalPadding

    Dim i As Integer
    For i = 1 To m_Stations.Count
        Dim sta As Station : Set sta = m_Stations(i)
        blockProxy.AddLineXYZXYZ sta.StationOffset, 0, 0, sta.StationOffset, HalfBreadth, 0
    Next i
    For i = 1 To m_SheerPlane.Count
        Dim sheerPlane As Double : sheerPlane = m_SheerPlane(i)
        If sheerPlane = 0 Or sheerPlane = HalfBreadth Then
            GoTo ContinueLoop
        End If
        blockProxy.AddLineXYZXYZ startLength, sheerPlane, 0, endLength, sheerPlane, 0
ContinueLoop:
    Next i
    ' BaseLine
    blockProxy.AddLineXYZXYZ startLength, 0, 0, endLength, 0, 0
    ' TopLine
    blockProxy.AddLineXYZXYZ startLength, HalfBreadth, 0, endLength, HalfBreadth, 0
End Sub

