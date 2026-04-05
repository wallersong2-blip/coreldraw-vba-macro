Attribute VB_Name = "modPlacementFlow"
Option Explicit

' Fluxo principal para posicionar o máximo de peças selecionadas em um retângulo-alvo.

Public Sub RunPlacementFlow()
    On Error GoTo HandleFail

    EnsureDocumentReady

    Dim oldUnit As cdrUnit
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter

    Dim sr As ShapeRange
    Dim targetShape As Shape
    Dim targetRect As RectMM
    Dim pieces As Collection
    Dim occupiedRects As New Collection
    Dim placedCount As Long
    Dim skippedCount As Long

    Set sr = GetSelectionOrFail()
    Set targetShape = FindTargetRectangle(sr)
    targetRect = RectFromShape(targetShape)

    Set pieces = CollectPieces(sr, targetShape)
    Set pieces = SortPiecesByAreaDescending(pieces)

    PlacePieceCollection pieces, targetRect, occupiedRects, placedCount, skippedCount

    ActiveDocument.Unit = oldUnit
    MsgBox "Concluído." & vbCrLf & _
           "Peças posicionadas: " & CStr(placedCount) & vbCrLf & _
           "Sem espaço: " & CStr(skippedCount), vbInformation, "Distribuição por Bounding Box"
    Exit Sub

HandleFail:
    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    MsgBox "Erro: " & Err.Description, vbExclamation, "RunPlacementFlow"
End Sub

Public Sub ExecutePlaceSelectedShapes()
    ' Ponto de entrada simples para executar a macro de fase única.
    RunPlacementFlow
End Sub

Public Sub ExecutePlaceSelectedShapesTwoPhase()
    ' Ponto de entrada para teste em duas fases sem layer fixa.
    ' Selecione alvo + todas as peças e marque as da fase 2 com Name começando em "F2_".
    On Error GoTo HandleFail

    EnsureDocumentReady

    Dim oldUnit As cdrUnit
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter

    Dim sr As ShapeRange
    Dim targetShape As Shape
    Dim targetRect As RectMM

    Dim phase1Pieces As Collection
    Dim phase2Pieces As Collection
    Dim occupiedRects As New Collection

    Dim placed1 As Long, skipped1 As Long
    Dim placed2 As Long, skipped2 As Long

    Set sr = GetSelectionOrFail()
    Set targetShape = FindTargetRectangle(sr)
    targetRect = RectFromShape(targetShape)

    SplitPiecesByPhaseFromSelection sr, targetShape, phase1Pieces, phase2Pieces

    Set phase1Pieces = SortPiecesByAreaDescending(phase1Pieces)
    Set phase2Pieces = SortPiecesByAreaDescending(phase2Pieces)

    PlacePieceCollection phase1Pieces, targetRect, occupiedRects, placed1, skipped1

    ' Fase 2 usa occupiedRects da fase 1 como obstáculo fixo.
    PlacePieceCollection phase2Pieces, targetRect, occupiedRects, placed2, skipped2

    ActiveDocument.Unit = oldUnit
    MsgBox "Concluído (2 fases)." & vbCrLf & _
           "Fase 1 - posicionadas: " & CStr(placed1) & " | sem espaço: " & CStr(skipped1) & vbCrLf & _
           "Fase 2 - posicionadas: " & CStr(placed2) & " | sem espaço: " & CStr(skipped2), _
           vbInformation, "Distribuição 2 Fases"
    Exit Sub

HandleFail:
    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    MsgBox "Erro: " & Err.Description, vbExclamation, "ExecutePlaceSelectedShapesTwoPhase"
End Sub

Private Sub PlacePieceCollection(ByVal pieces As Collection, _
                                 ByVal targetRect As RectMM, _
                                 ByRef occupiedRects As Collection, _
                                 ByRef placedCount As Long, _
                                 ByRef skippedCount As Long)
    Dim i As Long
    Dim shp As Shape
    Dim best As PlacementCandidate
    Dim baseRotation As Double

    For i = 1 To pieces.Count
        Set shp = pieces(i)
        baseRotation = shp.RotationAngle

        best = EvaluateBestCandidate(shp, targetRect, occupiedRects, GRID_STEP_FINE_MM, MIN_GAP_MM, baseRotation)

        If best.IsValid Then
            shp.Rotate best.RotationDeg - shp.RotationAngle
            MoveShapeCenterTo shp, best.X, best.Y
            occupiedRects.Add RectToArray(RectFromShape(shp))
            placedCount = placedCount + 1
        Else
            shp.Rotate baseRotation - shp.RotationAngle
            skippedCount = skippedCount + 1
        End If
    Next i
End Sub

Private Function SortPiecesByAreaDescending(ByVal pieces As Collection) As Collection
    ' Ordena por área de bounding box (maiores primeiro).
    Dim sorted As New Collection
    Dim inserted As Boolean
    Dim p As Shape
    Dim i As Long
    Dim pos As Long

    For i = 1 To pieces.Count
        Set p = pieces(i)
        inserted = False

        If sorted.Count = 0 Then
            sorted.Add p
            inserted = True
        Else
            For pos = 1 To sorted.Count
                If RectArea(RectFromShape(p)) > RectArea(RectFromShape(sorted(pos))) Then
                    sorted.Add p, Before:=pos
                    inserted = True
                    Exit For
                End If
            Next pos

            If Not inserted Then
                sorted.Add p
                inserted = True
            End If
        End If
    Next i

    Set SortPiecesByAreaDescending = sorted
End Function
