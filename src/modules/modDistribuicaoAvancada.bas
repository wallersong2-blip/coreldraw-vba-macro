Attribute VB_Name = "modDistribuicaoAvancada"
Option Explicit

Private Const MM_TO_INCH As Double = 1# / 25.4
Private Const SAMPLE_COUNT As Long = 320

' Folga mínima fixa entre peças = 0,5 mm
Private Const MIN_CLEAR_MM As Double = 0.5

' Compensação para reduzir o conservadorismo do perfil
' Quanto maior, mais ele aproxima as linhas.
' Se começar a encostar demais, reduza para 0.20 ou 0.15
Private Const PROFILE_RELIEF_MM As Double = 0.28

' Passo para tentativa de preenchimento de sobras
Private Const FILL_STEP_MM As Double = 0.5

Private Type TRect
    l As Double
    t As Double
    r As Double
    b As Double
End Type

Sub AbrirDistribuidor()
    frmDistribuir.Show
End Sub

Sub ExecutarDistribuicao()

    Dim sr As ShapeRange
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim base As Shape
    Dim area As Shape

    Dim Espacamento As Double
    Dim Margem As Double
    Dim passoManual As Double
    Dim maxCopias As Long
    Dim usarAlternancia As Boolean
    Dim centralizar As Boolean

    Dim larguraBase As Double
    Dim alturaBase As Double
    Dim larguraUtil As Double
    Dim alturaUtil As Double

    Dim leftLimite As Double
    Dim rightLimite As Double
    Dim topLimite As Double
    Dim bottomLimite As Double

    Dim topN() As Double
    Dim botN() As Double
    Dim topR() As Double
    Dim botR() As Double

    Dim melhorPitch As Double
    Dim melhorOffset As Double
    Dim melhorAlternancia As Boolean
    Dim melhorQtd As Long
    Dim melhorSobra As Double
    Dim melhorCompactacao As Double

    Dim blocked As Collection
    Dim totalCriadas As Long

    If ActiveSelectionRange.Count <> 2 Then
        MsgBox "Selecione exatamente 2 objetos: o nome e o retângulo."
        Exit Sub
    End If

    Set sr = ActiveSelectionRange
    Set shp1 = sr.Shapes(1)
    Set shp2 = sr.Shapes(2)

    If shp1.Type = cdrRectangleShape Then
        Set area = shp1
        Set base = shp2
    ElseIf shp2.Type = cdrRectangleShape Then
        Set area = shp2
        Set base = shp1
    Else
        MsgBox "Um dos objetos precisa ser um retângulo."
        Exit Sub
    End If

    If Trim$(frmDistribuir.txtQuantidade.Value) = "" Then
        maxCopias = 999999
    Else
        maxCopias = CLng(frmDistribuir.txtQuantidade.Value)
        If maxCopias < 1 Then
            MsgBox "Quantidade máxima inválida."
            Exit Sub
        End If
    End If

    ' A distância horizontal agora respeita no mínimo 0,5 mm
    Espacamento = CDbl(frmDistribuir.txtEspacamento.Value) * MM_TO_INCH
    If Espacamento < (MIN_CLEAR_MM * MM_TO_INCH) Then
        Espacamento = MIN_CLEAR_MM * MM_TO_INCH
    End If

    Margem = CDbl(frmDistribuir.txtMargem.Value) * MM_TO_INCH

    If Trim$(frmDistribuir.txtPassoVertical.Value) = "" Then
        passoManual = 0
    Else
        passoManual = CDbl(frmDistribuir.txtPassoVertical.Value) * MM_TO_INCH
    End If

    usarAlternancia = frmDistribuir.chkRotacao.Value
    centralizar = frmDistribuir.chkCentralizar.Value

    larguraBase = base.SizeWidth
    alturaBase = base.SizeHeight

    If larguraBase <= 0 Or alturaBase <= 0 Then
        MsgBox "Objeto base inválido."
        Exit Sub
    End If

    leftLimite = area.leftX + Margem
    rightLimite = area.rightX - Margem
    topLimite = area.topY - Margem
    bottomLimite = area.bottomY + Margem

    larguraUtil = rightLimite - leftLimite
    alturaUtil = topLimite - bottomLimite

    If larguraUtil <= 0 Or alturaUtil <= 0 Then
        MsgBox "A área útil do retângulo ficou inválida."
        Exit Sub
    End If

    If larguraUtil < larguraBase Or alturaUtil < alturaBase Then
        MsgBox "O nome não cabe no retângulo."
        Exit Sub
    End If

    Set blocked = New Collection
    ColetarObstaculos area, base, leftLimite, rightLimite, topLimite, bottomLimite, blocked

    BuildProfiles base, SAMPLE_COUNT, topN, botN, larguraBase, alturaBase
    BuildRot180Profiles topN, botN, topR, botR, alturaBase

    melhorQtd = 0
    melhorPitch = 0
    melhorOffset = 0
    melhorAlternancia = False
    melhorSobra = 999999
    melhorCompactacao = -999999

    BuscarMelhorConfiguracao _
        larguraBase, alturaBase, larguraUtil, alturaUtil, Espacamento, passoManual, _
        topN, botN, topR, botR, usarAlternancia, blocked, leftLimite, bottomLimite, _
        melhorPitch, melhorOffset, melhorAlternancia, melhorQtd, melhorSobra, melhorCompactacao

    If melhorQtd < 1 Then
        MsgBox "Nada coube dentro do retângulo."
        Exit Sub
    End If

    totalCriadas = DistribuirMelhorConfiguracao( _
        base, larguraBase, alturaBase, _
        leftLimite, rightLimite, topLimite, bottomLimite, _
        Espacamento, melhorPitch, melhorOffset, melhorAlternancia, _
        maxCopias, centralizar, blocked)

    ' Fase de preenchimento de sobras com busca por grade fina sem sobrepor.
    If totalCriadas < maxCopias Then
        totalCriadas = totalCriadas + PreencherSobras( _
            base, leftLimite, rightLimite, topLimite, bottomLimite, _
            maxCopias - totalCriadas, blocked)
    End If

    MsgBox totalCriadas & " cópias criadas!"

End Sub

Private Sub BuscarMelhorConfiguracao( _
    ByVal larguraBase As Double, _
    ByVal alturaBase As Double, _
    ByVal larguraUtil As Double, _
    ByVal alturaUtil As Double, _
    ByVal Espacamento As Double, _
    ByVal passoManual As Double, _
    ByRef topN() As Double, _
    ByRef botN() As Double, _
    ByRef topR() As Double, _
    ByRef botR() As Double, _
    ByVal permitirAlternancia As Boolean, _
    ByVal blocked As Collection, _
    ByVal leftLimite As Double, _
    ByVal bottomLimite As Double, _
    ByRef melhorPitch As Double, _
    ByRef melhorOffset As Double, _
    ByRef melhorAlternancia As Boolean, _
    ByRef melhorQtd As Long, _
    ByRef melhorSobra As Double, _
    ByRef melhorCompactacao As Double)

    Dim offset As Double
    Dim stepOffset As Double
    Dim pitch As Double
    Dim qtd As Long
    Dim sobra As Double
    Dim safePitch As Double
    Dim maxOffset As Double
    Dim clearV As Double
    Dim compactacao As Double

    clearV = MIN_CLEAR_MM * MM_TO_INCH

    melhorQtd = 0
    melhorPitch = 0
    melhorOffset = 0
    melhorAlternancia = False
    melhorSobra = 999999
    melhorCompactacao = -999999

    ' cenário sem alternância
    safePitch = RequiredPitch(topN, botN, topN, botN, larguraBase, 0, clearV)
    pitch = MaxD(ApplyProfileRelief(safePitch), passoManual)

    qtd = SimularQuantidade(larguraBase, alturaBase, larguraUtil, alturaUtil, Espacamento, pitch, 0, False, blocked, leftLimite, bottomLimite)
    sobra = CalcularSobraHorizontal(larguraBase, larguraUtil, Espacamento, 0, False, alturaUtil, alturaBase, pitch)
    compactacao = CalcularCompactacao(alturaUtil, alturaBase, pitch, False, 0, larguraUtil, Espacamento, larguraBase)

    AtualizarMelhor _
        qtd, pitch, 0, False, sobra, compactacao, _
        melhorQtd, melhorPitch, melhorOffset, melhorAlternancia, melhorSobra, melhorCompactacao

    If permitirAlternancia Then

        ' Busca ampla e fina
        maxOffset = larguraBase * 0.98
        stepOffset = larguraBase / 80#
        If stepOffset <= 0 Then stepOffset = 0.01

        offset = 0

        Do While offset <= maxOffset + 0.000001

            safePitch = RequiredPitch(topN, botN, topR, botR, larguraBase, offset, clearV)
            pitch = MaxD(ApplyProfileRelief(safePitch), passoManual)

            qtd = SimularQuantidade(larguraBase, alturaBase, larguraUtil, alturaUtil, Espacamento, pitch, offset, True, blocked, leftLimite, bottomLimite)
            sobra = CalcularSobraHorizontal(larguraBase, larguraUtil, Espacamento, offset, True, alturaUtil, alturaBase, pitch)
            compactacao = CalcularCompactacao(alturaUtil, alturaBase, pitch, True, offset, larguraUtil, Espacamento, larguraBase)

            AtualizarMelhor _
                qtd, pitch, offset, True, sobra, compactacao, _
                melhorQtd, melhorPitch, melhorOffset, melhorAlternancia, melhorSobra, melhorCompactacao

            offset = offset + stepOffset
        Loop
    End If

End Sub

Private Sub AtualizarMelhor( _
    ByVal qtd As Long, _
    ByVal pitch As Double, _
    ByVal offset As Double, _
    ByVal alternancia As Boolean, _
    ByVal sobra As Double, _
    ByVal compactacao As Double, _
    ByRef melhorQtd As Long, _
    ByRef melhorPitch As Double, _
    ByRef melhorOffset As Double, _
    ByRef melhorAlternancia As Boolean, _
    ByRef melhorSobra As Double, _
    ByRef melhorCompactacao As Double)

    If qtd > melhorQtd Then
        melhorQtd = qtd
        melhorPitch = pitch
        melhorOffset = offset
        melhorAlternancia = alternancia
        melhorSobra = sobra
        melhorCompactacao = compactacao
        Exit Sub
    End If

    If qtd = melhorQtd And qtd > 0 Then

        ' prioridade forte para alternância quando empata
        If alternancia And Not melhorAlternancia Then
            melhorQtd = qtd
            melhorPitch = pitch
            melhorOffset = offset
            melhorAlternancia = alternancia
            melhorSobra = sobra
            melhorCompactacao = compactacao
            Exit Sub
        End If

        If alternancia = melhorAlternancia Then

            ' prioridade 1: maior compactação
            If compactacao > melhorCompactacao Then
                melhorQtd = qtd
                melhorPitch = pitch
                melhorOffset = offset
                melhorAlternancia = alternancia
                melhorSobra = sobra
                melhorCompactacao = compactacao
                Exit Sub
            End If

            ' prioridade 2: menor pitch
            If Abs(pitch - melhorPitch) > 0.000001 Then
                If pitch < melhorPitch Or melhorPitch = 0 Then
                    melhorQtd = qtd
                    melhorPitch = pitch
                    melhorOffset = offset
                    melhorAlternancia = alternancia
                    melhorSobra = sobra
                    melhorCompactacao = compactacao
                    Exit Sub
                End If
            End If

            ' prioridade 3: menor sobra horizontal
            If sobra < melhorSobra Then
                melhorQtd = qtd
                melhorPitch = pitch
                melhorOffset = offset
                melhorAlternancia = alternancia
                melhorSobra = sobra
                melhorCompactacao = compactacao
                Exit Sub
            End If
        End If
    End If

End Sub

Private Function ApplyProfileRelief(ByVal safePitch As Double) As Double
    Dim alivio As Double

    alivio = PROFILE_RELIEF_MM * MM_TO_INCH

    ApplyProfileRelief = safePitch - alivio
    If ApplyProfileRelief < 0 Then ApplyProfileRelief = 0
End Function

Private Function SimularQuantidade( _
    ByVal larguraBase As Double, _
    ByVal alturaBase As Double, _
    ByVal larguraUtil As Double, _
    ByVal alturaUtil As Double, _
    ByVal Espacamento As Double, _
    ByVal pitch As Double, _
    ByVal offsetOdd As Double, _
    ByVal alternancia As Boolean, _
    ByVal blocked As Collection, _
    ByVal leftLimite As Double, _
    ByVal bottomLimite As Double) As Long

    Dim rows As Long
    Dim r As Long
    Dim c As Long
    Dim colsRow As Long
    Dim rowOffset As Double
    Dim startX As Double
    Dim atualX As Double
    Dim atualY As Double
    Dim qtd As Long

    If pitch <= 0 Then
        SimularQuantidade = 0
        Exit Function
    End If

    If alturaUtil < alturaBase Then
        SimularQuantidade = 0
        Exit Function
    End If

    rows = Int((alturaUtil - alturaBase) / pitch) + 1
    If rows < 1 Then
        SimularQuantidade = 0
        Exit Function
    End If

    qtd = 0

    For r = 0 To rows - 1
        atualY = bottomLimite + alturaBase + (r * pitch)

        If alternancia And ((r Mod 2) = 1) Then
            rowOffset = offsetOdd
        Else
            rowOffset = 0
        End If

        colsRow = FitCount(larguraUtil, larguraBase, Espacamento, rowOffset)
        If colsRow > 0 Then
            startX = leftLimite + rowOffset
            For c = 0 To colsRow - 1
                atualX = startX + (c * (larguraBase + Espacamento))
                If Not RetanguloColide(atualX, atualY, larguraBase, alturaBase, blocked, 0) Then
                    qtd = qtd + 1
                End If
            Next c
        End If
    Next r

    SimularQuantidade = qtd

End Function

Private Function CalcularCompactacao( _
    ByVal alturaUtil As Double, _
    ByVal alturaBase As Double, _
    ByVal pitch As Double, _
    ByVal alternancia As Boolean, _
    ByVal offsetOdd As Double, _
    ByVal larguraUtil As Double, _
    ByVal Espacamento As Double, _
    ByVal larguraBase As Double) As Double

    Dim rows As Long
    Dim evenCols As Long
    Dim oddCols As Long
    Dim ocupacaoH As Double
    Dim ganhoOffset As Double

    If pitch <= 0 Then
        CalcularCompactacao = -999999
        Exit Function
    End If

    rows = Int((alturaUtil - alturaBase) / pitch) + 1
    If rows < 1 Then
        CalcularCompactacao = -999999
        Exit Function
    End If

    evenCols = FitCount(larguraUtil, larguraBase, Espacamento, 0)

    If alternancia Then
        oddCols = FitCount(larguraUtil, larguraBase, Espacamento, offsetOdd)
        ganhoOffset = offsetOdd / larguraBase
    Else
        oddCols = evenCols
        ganhoOffset = 0
    End If

    ocupacaoH = (evenCols + oddCols) / 2#

    ' Quanto maior, melhor
    CalcularCompactacao = (rows * 1000#) - (pitch * 100#) + ocupacaoH + ganhoOffset
End Function

Private Function CalcularSobraHorizontal( _
    ByVal larguraBase As Double, _
    ByVal larguraUtil As Double, _
    ByVal Espacamento As Double, _
    ByVal offsetOdd As Double, _
    ByVal alternancia As Boolean, _
    ByVal alturaUtil As Double, _
    ByVal alturaBase As Double, _
    ByVal pitch As Double) As Double

    Dim rows As Long
    Dim r As Long
    Dim colsRow As Long
    Dim rowOffset As Double
    Dim ocupada As Double
    Dim sobra As Double
    Dim soma As Double

    If pitch <= 0 Then
        CalcularSobraHorizontal = 999999
        Exit Function
    End If

    rows = Int((alturaUtil - alturaBase) / pitch) + 1
    If rows < 1 Then rows = 1

    soma = 0

    For r = 0 To rows - 1
        If alternancia And ((r Mod 2) = 1) Then
            rowOffset = offsetOdd
        Else
            rowOffset = 0
        End If

        colsRow = FitCount(larguraUtil, larguraBase, Espacamento, rowOffset)

        If colsRow > 0 Then
            ocupada = rowOffset + (colsRow * larguraBase) + ((colsRow - 1) * Espacamento)
            sobra = larguraUtil - ocupada
            soma = soma + sobra
        Else
            soma = soma + larguraUtil
        End If
    Next r

    CalcularSobraHorizontal = soma

End Function

Private Function FitCount( _
    ByVal larguraUtil As Double, _
    ByVal larguraBase As Double, _
    ByVal Espacamento As Double, _
    ByVal offset As Double) As Long

    Dim restante As Double

    restante = larguraUtil - offset
    If restante < larguraBase Then
        FitCount = 0
    Else
        FitCount = Int((restante + Espacamento) / (larguraBase + Espacamento))
    End If

End Function

Private Function DistribuirMelhorConfiguracao( _
    ByVal base As Shape, _
    ByVal larguraBase As Double, _
    ByVal alturaBase As Double, _
    ByVal leftLimite As Double, _
    ByVal rightLimite As Double, _
    ByVal topLimite As Double, _
    ByVal bottomLimite As Double, _
    ByVal Espacamento As Double, _
    ByVal pitch As Double, _
    ByVal offsetOdd As Double, _
    ByVal alternancia As Boolean, _
    ByVal maxCopias As Long, _
    ByVal centralizar As Boolean, _
    ByVal blocked As Collection) As Long

    Dim larguraUtil As Double
    Dim alturaUtil As Double
    Dim rows As Long
    Dim r As Long
    Dim c As Long
    Dim i As Long

    Dim atualY As Double
    Dim atualX As Double
    Dim rowOffset As Double
    Dim startX As Double
    Dim colsRow As Long
    Dim occupiedWidth As Double

    Dim copia As Shape

    larguraUtil = rightLimite - leftLimite
    alturaUtil = topLimite - bottomLimite

    rows = Int((alturaUtil - alturaBase) / pitch) + 1
    If rows < 1 Then
        DistribuirMelhorConfiguracao = 0
        Exit Function
    End If

    i = 0

    For r = 0 To rows - 1

        atualY = topLimite - (r * pitch)

        If alternancia And ((r Mod 2) = 1) Then
            rowOffset = offsetOdd
        Else
            rowOffset = 0
        End If

        colsRow = FitCount(larguraUtil, larguraBase, Espacamento, rowOffset)
        If colsRow < 1 Then GoTo ProximaLinha

        If centralizar Then
            occupiedWidth = (colsRow * larguraBase) + ((colsRow - 1) * Espacamento)
            startX = leftLimite + rowOffset + ((larguraUtil - rowOffset - occupiedWidth) / 2#)
        Else
            startX = leftLimite + rowOffset
        End If

        For c = 0 To colsRow - 1
            atualX = startX + (c * (larguraBase + Espacamento))

            If Not RetanguloColide(atualX, atualY, larguraBase, alturaBase, blocked, MIN_CLEAR_MM * MM_TO_INCH / 2#) Then
                Set copia = base.Duplicate

                If alternancia And ((r Mod 2) = 1) Then
                    copia.Rotate 180
                End If

                copia.Move atualX - copia.leftX, atualY - copia.topY
                blocked.Add MakeRect(copia.leftX, copia.topY, copia.rightX, copia.bottomY)

                i = i + 1
                If i >= maxCopias Then Exit For
            End If
        Next c

        If i >= maxCopias Then Exit For

ProximaLinha:
    Next r

    DistribuirMelhorConfiguracao = i

End Function

Private Function PreencherSobras( _
    ByVal base As Shape, _
    ByVal leftLimite As Double, _
    ByVal rightLimite As Double, _
    ByVal topLimite As Double, _
    ByVal bottomLimite As Double, _
    ByVal maxAdicionar As Long, _
    ByVal blocked As Collection) As Long

    Dim stepXY As Double
    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double
    Dim rot As Long
    Dim added As Long
    Dim copia As Shape
    Dim oldRot As Double

    oldRot = base.RotationAngle

    stepXY = FILL_STEP_MM * MM_TO_INCH
    If stepXY <= 0 Then stepXY = MIN_CLEAR_MM * MM_TO_INCH

    added = 0

    ' Busca gulosa de cima para baixo e esquerda para direita,
    ' testando 0° e 180° para preencher sobras sem sobreposição.
    y = topLimite
    Do While y >= bottomLimite
        x = leftLimite
        Do While x <= rightLimite
            For rot = 0 To 1
                If rot = 0 Then
                    base.Rotate -base.RotationAngle
                Else
                    base.Rotate 180 - base.RotationAngle
                End If

                w = base.SizeWidth
                h = base.SizeHeight

                If (x + w <= rightLimite) And (y - h >= bottomLimite) Then
                    If Not RetanguloColide(x, y, w, h, blocked, MIN_CLEAR_MM * MM_TO_INCH / 2#) Then
                        Set copia = base.Duplicate
                        If rot = 1 Then copia.Rotate 180
                        copia.Move x - copia.leftX, y - copia.topY

                        blocked.Add MakeRect(copia.leftX, copia.topY, copia.rightX, copia.bottomY)
                        added = added + 1
                        Exit For
                    End If
                End If
            Next rot

            If added >= maxAdicionar Then
                PreencherSobras = added
                Exit Function
            End If

            x = x + stepXY
        Loop
        y = y - stepXY
    Loop

    base.Rotate oldRot - base.RotationAngle
    PreencherSobras = added
End Function

Private Sub ColetarObstaculos( _
    ByVal area As Shape, _
    ByVal base As Shape, _
    ByVal leftLimite As Double, _
    ByVal rightLimite As Double, _
    ByVal topLimite As Double, _
    ByVal bottomLimite As Double, _
    ByVal blocked As Collection)

    Dim shp As Shape
    Dim r As TRect
    Dim areaID As Long
    Dim baseID As Long

    areaID = area.StaticID
    baseID = base.StaticID

    For Each shp In ActivePage.Shapes
        If shp.StaticID <> areaID And shp.StaticID <> baseID Then
            r = MakeRect(shp.leftX, shp.topY, shp.rightX, shp.bottomY)
            If RectIntersects(r, MakeRect(leftLimite, topLimite, rightLimite, bottomLimite)) Then
                blocked.Add r
            End If
        End If
    Next shp
End Sub

Private Function RetanguloColide( _
    ByVal leftX As Double, _
    ByVal topY As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal blocked As Collection, _
    ByVal halfGap As Double) As Boolean

    Dim cand As TRect
    Dim occ As TRect
    Dim i As Long

    cand = MakeRect(leftX - halfGap, topY + halfGap, leftX + w + halfGap, topY - h - halfGap)

    For i = 1 To blocked.Count
        occ = blocked(i)
        If RectIntersects(cand, occ) Then
            RetanguloColide = True
            Exit Function
        End If
    Next i

    RetanguloColide = False
End Function

Private Function MakeRect(ByVal leftX As Double, ByVal topY As Double, ByVal rightX As Double, ByVal bottomY As Double) As TRect
    Dim r As TRect
    r.l = leftX
    r.t = topY
    r.r = rightX
    r.b = bottomY
    MakeRect = NormalizeRect(r)
End Function

Private Function NormalizeRect(ByVal r As TRect) As TRect
    Dim n As TRect

    If r.l <= r.r Then
        n.l = r.l
        n.r = r.r
    Else
        n.l = r.r
        n.r = r.l
    End If

    If r.t >= r.b Then
        n.t = r.t
        n.b = r.b
    Else
        n.t = r.b
        n.b = r.t
    End If

    NormalizeRect = n
End Function

Private Function RectIntersects(ByVal a As TRect, ByVal b As TRect) As Boolean
    Dim nA As TRect
    Dim nB As TRect

    nA = NormalizeRect(a)
    nB = NormalizeRect(b)

    RectIntersects = Not (nA.r <= nB.l Or nA.l >= nB.r Or nA.t <= nB.b Or nA.b >= nB.t)
End Function

Private Function RequiredPitch( _
    ByRef topA() As Double, _
    ByRef botA() As Double, _
    ByRef topB() As Double, _
    ByRef botB() As Double, _
    ByVal larguraBase As Double, _
    ByVal offsetB As Double, _
    ByVal clearV As Double) As Double

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim dx As Double
    Dim xA As Double
    Dim xB As Double
    Dim req As Double
    Dim cand As Double

    n = UBound(topA) - LBound(topA) + 1
    dx = larguraBase / n
    req = 0

    For i = 0 To n - 1
        If botA(i) >= topA(i) Then
            xA = (i + 0.5) * dx
            xB = xA - offsetB

            If xB >= 0 And xB < larguraBase Then
                j = Int(xB / dx)

                If botB(j) >= topB(j) Then
                    cand = botA(i) - topB(j) + clearV
                    If cand > req Then req = cand
                End If
            End If
        End If
    Next i

    If req < 0 Then req = 0
    RequiredPitch = req

End Function

Private Sub BuildRot180Profiles( _
    ByRef topN() As Double, _
    ByRef botN() As Double, _
    ByRef topR() As Double, _
    ByRef botR() As Double, _
    ByVal alturaBase As Double)

    Dim i As Long
    Dim j As Long
    Dim n As Long

    n = UBound(topN) - LBound(topN) + 1
    ReDim topR(0 To n - 1)
    ReDim botR(0 To n - 1)

    For i = 0 To n - 1
        j = n - 1 - i

        If botN(j) >= topN(j) Then
            topR(i) = alturaBase - botN(j)
            botR(i) = alturaBase - topN(j)
        Else
            topR(i) = alturaBase
            botR(i) = -1
        End If
    Next i

End Sub

Private Sub BuildProfiles( _
    ByVal base As Shape, _
    ByVal sampleCount As Long, _
    ByRef topProf() As Double, _
    ByRef botProf() As Double, _
    ByVal larguraBase As Double, _
    ByVal alturaBase As Double)

    Dim i As Long

    ReDim topProf(0 To sampleCount - 1)
    ReDim botProf(0 To sampleCount - 1)

    For i = 0 To sampleCount - 1
        topProf(i) = alturaBase
        botProf(i) = -1
    Next i

    AddShapeToProfileRecursive base, base.leftX, base.topY, sampleCount, topProf, botProf, larguraBase, alturaBase

End Sub

Private Sub AddShapeToProfileRecursive( _
    ByVal shp As Shape, _
    ByVal baseLeft As Double, _
    ByVal baseTop As Double, _
    ByVal sampleCount As Long, _
    ByRef topProf() As Double, _
    ByRef botProf() As Double, _
    ByVal larguraBase As Double, _
    ByVal alturaBase As Double)

    Dim child As Shape
    Dim childLeft As Double
    Dim childRight As Double
    Dim childTop As Double
    Dim childBottom As Double

    If shp.Type = cdrGroupShape Then
        For Each child In shp.Shapes
            AddShapeToProfileRecursive child, baseLeft, baseTop, sampleCount, topProf, botProf, larguraBase, alturaBase
        Next child
    Else
        childLeft = shp.leftX - baseLeft
        childRight = shp.rightX - baseLeft
        childTop = baseTop - shp.topY
        childBottom = baseTop - shp.bottomY

        AddBoxToProfile childLeft, childRight, childTop, childBottom, sampleCount, topProf, botProf, larguraBase, alturaBase
    End If

End Sub

Private Sub AddBoxToProfile( _
    ByVal leftX As Double, _
    ByVal rightX As Double, _
    ByVal topY As Double, _
    ByVal bottomY As Double, _
    ByVal sampleCount As Long, _
    ByRef topProf() As Double, _
    ByRef botProf() As Double, _
    ByVal larguraBase As Double, _
    ByVal alturaBase As Double)

    Dim i1 As Long
    Dim i2 As Long
    Dim i As Long
    Dim dx As Double

    If rightX <= 0 Or leftX >= larguraBase Then Exit Sub
    If bottomY <= 0 Or topY >= alturaBase Then Exit Sub

    If leftX < 0 Then leftX = 0
    If rightX > larguraBase Then rightX = larguraBase
    If topY < 0 Then topY = 0
    If bottomY > alturaBase Then bottomY = alturaBase

    dx = larguraBase / sampleCount

    i1 = Int(leftX / dx)
    i2 = Int((rightX - 0.0000001) / dx)

    If i1 < 0 Then i1 = 0
    If i2 > sampleCount - 1 Then i2 = sampleCount - 1

    For i = i1 To i2
        If topY < topProf(i) Then topProf(i) = topY
        If bottomY > botProf(i) Then botProf(i) = bottomY
    Next i

End Sub

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        MaxD = a
    Else
        MaxD = b
    End If
End Function
