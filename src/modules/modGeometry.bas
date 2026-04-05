Attribute VB_Name = "modGeometry"
Option Explicit

' Utilitários geométricos em mm usando bounding box do CorelDRAW.

Public Type RectMM
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
End Type

Public Function NormalizeRect(ByVal r As RectMM) As RectMM
    Dim outR As RectMM
    outR.Left = IIf(r.Left <= r.Right, r.Left, r.Right)
    outR.Right = IIf(r.Left <= r.Right, r.Right, r.Left)
    outR.Top = IIf(r.Top >= r.Bottom, r.Top, r.Bottom)
    outR.Bottom = IIf(r.Top >= r.Bottom, r.Bottom, r.Top)
    NormalizeRect = outR
End Function

Public Function RectWidth(ByVal r As RectMM) As Double
    Dim n As RectMM
    n = NormalizeRect(r)
    RectWidth = n.Right - n.Left
End Function

Public Function RectHeight(ByVal r As RectMM) As Double
    Dim n As RectMM
    n = NormalizeRect(r)
    RectHeight = n.Top - n.Bottom
End Function

Public Function RectArea(ByVal r As RectMM) As Double
    RectArea = RectWidth(r) * RectHeight(r)
End Function

Public Function InflateRect(ByVal r As RectMM, ByVal deltaMM As Double) As RectMM
    Dim n As RectMM
    n = NormalizeRect(r)
    n.Left = n.Left - deltaMM
    n.Top = n.Top + deltaMM
    n.Right = n.Right + deltaMM
    n.Bottom = n.Bottom - deltaMM
    InflateRect = n
End Function

Public Function Intersects(ByVal a As RectMM, ByVal b As RectMM) As Boolean
    Dim nA As RectMM, nB As RectMM
    nA = NormalizeRect(a)
    nB = NormalizeRect(b)

    Intersects = Not (nA.Right <= nB.Left Or nA.Left >= nB.Right Or nA.Top <= nB.Bottom Or nA.Bottom >= nB.Top)
End Function

Public Function ContainsRect(ByVal container As RectMM, ByVal inner As RectMM) As Boolean
    Dim c As RectMM, i As RectMM
    c = NormalizeRect(container)
    i = NormalizeRect(inner)

    ContainsRect = (i.Left >= c.Left And i.Right <= c.Right And i.Top <= c.Top And i.Bottom >= c.Bottom)
End Function

Public Function RectFromShape(ByVal shp As Shape) As RectMM
    Dim r As RectMM
    r.Left = shp.LeftX
    r.Top = shp.TopY
    r.Right = shp.RightX
    r.Bottom = shp.BottomY
    RectFromShape = NormalizeRect(r)
End Function

Public Sub MoveShapeCenterTo(ByVal shp As Shape, ByVal centerX As Double, ByVal centerY As Double)
    Dim dx As Double, dy As Double
    dx = centerX - shp.CenterX
    dy = centerY - shp.CenterY
    shp.Move dx, dy
End Sub

Public Function RectToArray(ByVal r As RectMM) As Variant
    Dim outA(1 To 4) As Double
    Dim n As RectMM
    n = NormalizeRect(r)
    outA(1) = n.Left
    outA(2) = n.Top
    outA(3) = n.Right
    outA(4) = n.Bottom
    RectToArray = outA
End Function

Public Function RectFromArray(ByVal values As Variant) As RectMM
    Dim r As RectMM
    r.Left = CDbl(values(1))
    r.Top = CDbl(values(2))
    r.Right = CDbl(values(3))
    r.Bottom = CDbl(values(4))
    RectFromArray = NormalizeRect(r)
End Function
