Attribute VB_Name = "modPackingEngine"
Option Explicit

' Motor de encaixe com prioridade para pontos candidatos inteligentes.
' Estratégia: tenta alinhamentos úteis com bordas do alvo e com peças já colocadas,
' e usa grid fino apenas como fallback.

Public Type PlacementCandidate
    X As Double
    Y As Double
    RotationDeg As Double
    Score As Double
    IsValid As Boolean
End Type

Private Function CandidateCollides(ByVal candidateRect As RectMM, ByVal occupiedRects As Collection, ByVal halfGap As Double) As Boolean
    Dim i As Long
    Dim inflatedCandidate As RectMM
    Dim occupied As RectMM

    inflatedCandidate = InflateRect(candidateRect, halfGap)

    For i = 1 To occupiedRects.Count
        occupied = RectFromArray(occupiedRects(i))
        If Intersects(inflatedCandidate, InflateRect(occupied, halfGap)) Then
            CandidateCollides = True
            Exit Function
        End If
    Next i

    CandidateCollides = False
End Function

Private Function AxisRectDistance(ByVal a As RectMM, ByVal b As RectMM) As Double
    Dim dx As Double, dy As Double

    If a.Right < b.Left Then
        dx = b.Left - a.Right
    ElseIf b.Right < a.Left Then
        dx = a.Left - b.Right
    Else
        dx = 0#
    End If

    If a.Top < b.Bottom Then
        dy = b.Bottom - a.Top
    ElseIf b.Top < a.Bottom Then
        dy = a.Bottom - b.Top
    Else
        dy = 0#
    End If

    AxisRectDistance = Sqr((dx * dx) + (dy * dy))
End Function

Private Function BoundaryMinDistance(ByVal targetRect As RectMM, ByVal candidateRect As RectMM) As Double
    Dim dL As Double, dR As Double, dT As Double, dB As Double
    dL = candidateRect.Left - targetRect.Left
    dR = targetRect.Right - candidateRect.Right
    dT = targetRect.Top - candidateRect.Top
    dB = candidateRect.Bottom - targetRect.Bottom

    BoundaryMinDistance = dL
    If dR < BoundaryMinDistance Then BoundaryMinDistance = dR
    If dT < BoundaryMinDistance Then BoundaryMinDistance = dT
    If dB < BoundaryMinDistance Then BoundaryMinDistance = dB
End Function

Private Function ComputeScore(ByVal targetRect As RectMM, _
                              ByVal candidateRect As RectMM, _
                              ByVal occupiedRects As Collection, _
                              ByVal minGapMM As Double) As Double
    ' Escolhe o melhor candidato priorizando compactação:
    ' 1) proximidade de peças já colocadas, 2) proximidade de bordas, 3) menos corredores vazios.
    Dim nearestOccupied As Double
    Dim i As Long
    Dim occ As RectMM
    Dim d As Double
    Dim touchCount As Long

    Dim dBoundary As Double
    Dim gapL As Double, gapR As Double, gapT As Double, gapB As Double
    Dim emptyCorridorPenalty As Double

    nearestOccupied = 1E+30

    For i = 1 To occupiedRects.Count
        occ = RectFromArray(occupiedRects(i))
        d = AxisRectDistance(candidateRect, occ)

        If d < nearestOccupied Then nearestOccupied = d
        If d <= (minGapMM * 1.2) Then touchCount = touchCount + 1
    Next i

    dBoundary = BoundaryMinDistance(targetRect, candidateRect)

    gapL = candidateRect.Left - targetRect.Left
    gapR = targetRect.Right - candidateRect.Right
    gapT = targetRect.Top - candidateRect.Top
    gapB = candidateRect.Bottom - targetRect.Bottom
    emptyCorridorPenalty = gapL + gapR + gapT + gapB

    If occupiedRects.Count = 0 Then
        ComputeScore = (SCORE_WEIGHT_BOUNDARY / (dBoundary + 0.05)) - (SCORE_WEIGHT_VOID * emptyCorridorPenalty)
    Else
        ComputeScore = _
            (SCORE_WEIGHT_CONTACT / (nearestOccupied + 0.05)) + _
            (SCORE_WEIGHT_BOUNDARY / (dBoundary + 0.05)) + _
            (SCORE_WEIGHT_TOUCH * CDbl(touchCount)) - _
            (SCORE_WEIGHT_VOID * emptyCorridorPenalty)
    End If
End Function

Private Sub TryUpdateBest(ByVal targetRect As RectMM, _
                          ByVal occupiedRects As Collection, _
                          ByVal minGapMM As Double, _
                          ByVal pieceW As Double, _
                          ByVal pieceH As Double, _
                          ByVal x As Double, _
                          ByVal y As Double, _
                          ByVal rotDeg As Double, _
                          ByRef best As PlacementCandidate)
    Dim candidateRect As RectMM
    Dim score As Double

    candidateRect.Left = x
    candidateRect.Right = x + pieceW
    candidateRect.Bottom = y
    candidateRect.Top = y + pieceH

    If ContainsRect(targetRect, candidateRect) Then
        If Not CandidateCollides(candidateRect, occupiedRects, minGapMM / 2#) Then
            score = ComputeScore(targetRect, candidateRect, occupiedRects, minGapMM)
            If (Not best.IsValid) Or (score > best.Score) Then
                best.IsValid = True
                best.X = x + (pieceW / 2#)
                best.Y = y + (pieceH / 2#)
                best.RotationDeg = rotDeg
                best.Score = score
            End If
        End If
    End If
End Sub

Private Sub AddUniqueValue(ByRef values As Collection, ByVal v As Double)
    Dim k As String
    k = CStr(CLng(Round(v * 1000#, 0)))

    On Error Resume Next
    values.Add v, k
    On Error GoTo 0
End Sub

Private Sub EvaluateSmartCandidates(ByVal targetRect As RectMM, _
                                    ByVal occupiedRects As Collection, _
                                    ByVal minGapMM As Double, _
                                    ByVal pieceW As Double, _
                                    ByVal pieceH As Double, _
                                    ByVal rotDeg As Double, _
                                    ByRef best As PlacementCandidate)
    Dim xs As New Collection
    Dim ys As New Collection
    Dim i As Long, ix As Long, iy As Long
    Dim occ As RectMM

    ' Bordas do retângulo-alvo.
    AddUniqueValue xs, targetRect.Left
    AddUniqueValue xs, targetRect.Right - pieceW
    AddUniqueValue ys, targetRect.Bottom
    AddUniqueValue ys, targetRect.Top - pieceH

    ' Alinhamentos úteis com peças já colocadas (encoste com gap mínimo).
    For i = 1 To occupiedRects.Count
        occ = RectFromArray(occupiedRects(i))

        ' Horizontal: peça nova à esquerda/direita da ocupada.
        AddUniqueValue xs, occ.Left - minGapMM - pieceW
        AddUniqueValue xs, occ.Right + minGapMM

        ' Vertical: peça nova abaixo/acima da ocupada.
        AddUniqueValue ys, occ.Bottom - minGapMM - pieceH
        AddUniqueValue ys, occ.Top + minGapMM

        ' Combinações extras de alinhamento de bordas para sobras irregulares.
        AddUniqueValue xs, occ.Left
        AddUniqueValue xs, occ.Right - pieceW
        AddUniqueValue ys, occ.Bottom
        AddUniqueValue ys, occ.Top - pieceH
    Next i

    ' Testa todas combinações relevantes X x Y.
    For ix = 1 To xs.Count
        For iy = 1 To ys.Count
            TryUpdateBest targetRect, occupiedRects, minGapMM, pieceW, pieceH, CDbl(xs(ix)), CDbl(ys(iy)), rotDeg, best
        Next iy
    Next ix
End Sub

Private Sub ScanGridFallback(ByVal targetRect As RectMM, _
                             ByVal searchRect As RectMM, _
                             ByVal occupiedRects As Collection, _
                             ByVal minGapMM As Double, _
                             ByVal pieceW As Double, _
                             ByVal pieceH As Double, _
                             ByVal stepMM As Double, _
                             ByVal rotDeg As Double, _
                             ByRef best As PlacementCandidate)
    Dim x As Double, y As Double

    y = searchRect.Bottom
    Do While y <= (searchRect.Top - pieceH)
        x = searchRect.Left
        Do While x <= (searchRect.Right - pieceW)
            TryUpdateBest targetRect, occupiedRects, minGapMM, pieceW, pieceH, x, y, rotDeg, best
            x = x + stepMM
        Loop
        y = y + stepMM
    Loop
End Sub

Public Function EvaluateBestCandidate(ByVal shp As Shape, _
                                      ByVal targetRect As RectMM, _
                                      ByVal occupiedRects As Collection, _
                                      ByVal gridStepMM As Double, _
                                      ByVal minGapMM As Double, _
                                      ByVal baseRotationDeg As Double) As PlacementCandidate

    Dim result As PlacementCandidate
    Dim rotations(0 To 1) As Double
    Dim rotIdx As Long

    Dim pieceRect As RectMM
    Dim pieceW As Double, pieceH As Double

    result.IsValid = False
    result.Score = -1E+30
    rotations(0) = baseRotationDeg
    rotations(1) = baseRotationDeg + 180#

    For rotIdx = LBound(rotations) To UBound(rotations)
        shp.Rotate rotations(rotIdx) - shp.RotationAngle
        pieceRect = RectFromShape(shp)
        pieceW = RectWidth(pieceRect)
        pieceH = RectHeight(pieceRect)

        ' Prioridade 1: pontos candidatos inteligentes (bordas + alinhamento com peças existentes).
        EvaluateSmartCandidates targetRect, occupiedRects, minGapMM, pieceW, pieceH, rotations(rotIdx), result

        ' Prioridade 2 (fallback): grid para não perder posições válidas.
        ScanGridFallback targetRect, targetRect, occupiedRects, minGapMM, pieceW, pieceH, gridStepMM, rotations(rotIdx), result

        ' Refino local perto do melhor ponto atual.
        If result.IsValid Then
            ScanGridFallback targetRect, BuildLocalRefineRect(targetRect, result.X, result.Y, gridStepMM), _
                             occupiedRects, minGapMM, pieceW, pieceH, GRID_STEP_MICRO_MM, rotations(rotIdx), result
        End If
    Next rotIdx

    shp.Rotate baseRotationDeg - shp.RotationAngle
    EvaluateBestCandidate = result
End Function

Private Function BuildLocalRefineRect(ByVal targetRect As RectMM, _
                                      ByVal centerX As Double, _
                                      ByVal centerY As Double, _
                                      ByVal baseStep As Double) As RectMM
    Dim outR As RectMM
    outR.Left = centerX - (3# * baseStep)
    outR.Right = centerX + (3# * baseStep)
    outR.Bottom = centerY - (3# * baseStep)
    outR.Top = centerY + (3# * baseStep)

    If outR.Left < targetRect.Left Then outR.Left = targetRect.Left
    If outR.Right > targetRect.Right Then outR.Right = targetRect.Right
    If outR.Bottom < targetRect.Bottom Then outR.Bottom = targetRect.Bottom
    If outR.Top > targetRect.Top Then outR.Top = targetRect.Top

    BuildLocalRefineRect = outR
End Function
