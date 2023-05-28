Option Explicit

Private m_Block As AcadBlock

Public Property Set Block(ByRef BlockRef As AcadBlock)
    Set m_Block = BlockRef
End Property

Public Function AddLine(StartPoint As Point3, EndPoint As Point3, Optional LayerName As String = "0")
    Set AddLine = m_Block.AddLine(StartPoint.ToArray(), EndPoint.ToArray())
    AddLine.Layer = LayerName
End Function

Public Function AddLineXYZXYZ(x1 as double, y1 as double, z1 as double, x2 as double, y2 as double, z2 as double, Optional LayerName As String = "0")
    Dim pt1 As New Point3
    Dim pt2 As New Point3
    pt1.XYZ x1,y1,z1
    pt2.XYZ x2,y2,z2
    Set AddLineXYZXYZ = AddLine(pt1, pt2)
    AddLineXYZXYZ.Layer = LayerName
End Function

Public Function AddSpline(Points As Point3Collection, StartTangent As Point3, EndTangent As Point3, Optional LayerName As String = "0")
    Set AddSpline = m_Block.AddSpline(Points.ToArray(), StartTangent.ToArray(), EndTangent.ToArray())
    AddSpline.Layer = LayerName
End Function