Attribute VB_Name = "modCorelAdapter"
Option Explicit

' Ponte com objetos do CorelDRAW usando somente VBA clássico.

Public Sub EnsureDocumentReady()
    If ActiveDocument Is Nothing Then
        Err.Raise vbObjectError + 7000, "modCorelAdapter", "Nenhum documento ativo no CorelDRAW."
    End If
End Sub

Public Function GetSelectionOrFail() As ShapeRange
    EnsureDocumentReady

    If ActiveSelectionRange Is Nothing Then
        Err.Raise vbObjectError + 7001, "modCorelAdapter", "Selecione um retângulo-alvo e as peças antes de executar."
    End If

    If ActiveSelectionRange.Count < 2 Then
        Err.Raise vbObjectError + 7002, "modCorelAdapter", "Selecione ao menos 1 retângulo-alvo e 1 peça."
    End If

    Set GetSelectionOrFail = ActiveSelectionRange
End Function

Public Function FindTargetRectangle(ByVal sr As ShapeRange) As Shape
    Dim i As Long
    Dim shp As Shape
    Dim bestArea As Double
    Dim thisArea As Double

    bestArea = -1#
    For i = 1 To sr.Count
        Set shp = sr(i)
        If shp.Type = cdrRectangleShape Then
            thisArea = RectArea(RectFromShape(shp))
            If thisArea > bestArea Then
                bestArea = thisArea
                Set FindTargetRectangle = shp
            End If
        End If
    Next i

    If FindTargetRectangle Is Nothing Then
        Err.Raise vbObjectError + 7003, "modCorelAdapter", "Nenhum retângulo selecionado foi encontrado."
    End If
End Function

Public Function CollectPieces(ByVal sr As ShapeRange, ByVal targetRect As Shape) As Collection
    Dim outC As New Collection
    Dim i As Long

    For i = 1 To sr.Count
        If sr(i).StaticID <> targetRect.StaticID Then
            outC.Add sr(i)
        End If
    Next i

    If outC.Count = 0 Then
        Err.Raise vbObjectError + 7004, "modCorelAdapter", "Nenhuma peça para posicionar foi encontrada na seleção."
    End If

    Set CollectPieces = outC
End Function

Public Sub SplitPiecesByPhaseFromSelection(ByVal sr As ShapeRange, _
                                           ByVal targetRect As Shape, _
                                           ByRef phase1 As Collection, _
                                           ByRef phase2 As Collection)
    ' Regra prática de uso diário:
    ' - fase 2 = shapes com Name iniciando em "F2_"
    ' - fase 1 = demais shapes da seleção (exceto retângulo-alvo)
    Dim i As Long
    Dim shp As Shape
    Dim shpName As String

    Set phase1 = New Collection
    Set phase2 = New Collection

    For i = 1 To sr.Count
        Set shp = sr(i)

        If shp.StaticID <> targetRect.StaticID Then
            shpName = UCase$(Trim$(shp.Name))

            If Left$(shpName, 3) = "F2_" Then
                phase2.Add shp
            Else
                phase1.Add shp
            End If
        End If
    Next i

    If phase1.Count = 0 Then
        Err.Raise vbObjectError + 7012, "modCorelAdapter", "Nenhuma peça de fase 1 encontrada."
    End If

    If phase2.Count = 0 Then
        Err.Raise vbObjectError + 7013, "modCorelAdapter", "Nenhuma peça de fase 2 encontrada. Use prefixo F2_ no Name dos shapes da fase 2."
    End If
End Sub
