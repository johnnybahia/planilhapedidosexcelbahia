Attribute VB_Name = "Modulo9"
Option Explicit

' ==========================================================================
' Cadastro de PONTA PERSONALIZADA.
' - Aba ITENS PONTA PERSONALIZADA (coluna A): itens-gatilho, um por celula.
' - Aba CADASTROS PONTA PERSONALIZADA: A = codigo do cliente (coluna D do
'   RELATORIO), B = descricao da ponta, C = "SIM" para nao perguntar mais.
' Chamado no fim do GerarRelatorio (Modulo3).
' ==========================================================================

Private Const ABA_CADASTRO As String = "CADASTROS PONTA PERSONALIZADA"
Private Const ABA_ITENS    As String = "ITENS PONTA PERSONALIZADA"
Private Const CAB_CODIGO   As String = "CODIGO CLIENTE"
Private Const CAB_DESC     As String = "DESCRICAO PONTA PERSONALIZADA"
Private Const CAB_IGNORAR  As String = "NAO PERGUNTAR MAIS"
Private Const CAB_ITENS    As String = "ITENS QUE ABREM O CADASTRO"
Private Const ITEM_EXEMPLO As String = "M13769"
Private Const MAX_LINHAS_AVISO As Long = 20

Public Sub ProcessarPontaPersonalizada(ByRef relatorio_ws As Worksheet, ByVal last_row As Long)
    Dim wsCad As Worksheet
    Dim dictCad As Object, dictGat As Object, dictPend As Object
    Dim gatilhos As Variant, dados As Variant, info As Variant
    Dim chavePend As Variant
    Dim n As Long, i As Long
    Dim cod As String, chave As String, textoF As String, desc As String, descricao As String
    Dim linhasSemCodigo As String, qtdSemCodigo As Long
    Dim resp As VbMsgBoxResult

    If last_row < 4 Then Exit Sub

    Set wsCad = AbaCadastro()
    Set dictCad = CarregarCadastro(wsCad)
    Set dictGat = CarregarGatilhos()
    gatilhos = dictGat.Keys
    Set dictPend = CreateObject("Scripting.Dictionary")

    ' Colunas D, E, F do RELATORIO em memoria (3 colunas = sempre matriz 2D)
    dados = relatorio_ws.Range(relatorio_ws.Cells(4, 4), relatorio_ws.Cells(last_row, 6)).Value
    n = UBound(dados, 1)

    ' --- FASE 1: detecta gatilhos na coluna F e levanta codigos sem cadastro ---
    If dictGat.Count > 0 Then
        For i = 1 To n
            textoF = TextoCelula(dados(i, 3))
            If textoF <> "" Then
                If ContemGatilho(textoF, gatilhos) Then
                    cod = Trim(TextoCelula(dados(i, 1)))
                    If cod = "" Then
                        qtdSemCodigo = qtdSemCodigo + 1
                        If qtdSemCodigo <= MAX_LINHAS_AVISO Then linhasSemCodigo = linhasSemCodigo & (i + 3) & " "
                    Else
                        chave = UCase(cod)
                        If Not dictPend.Exists(chave) Then
                            If dictCad.Exists(chave) Then
                                info = dictCad(chave)
                                If (Not info(2)) And info(1) = "" Then dictPend.Add chave, Array(cod, textoF)
                            Else
                                dictPend.Add chave, Array(cod, textoF)
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If

    ' Uma pergunta por codigo (nao por linha)
    For Each chavePend In dictPend.Keys
        info = dictPend(chavePend)
        resp = MsgBox("Item com ponta personalizada detectado." & vbCrLf & vbCrLf & _
            "Codigo do cliente: " & info(0) & vbCrLf & _
            "Descricao: " & Left(info(1), 200) & vbCrLf & vbCrLf & _
            "Sim = cadastrar a ponta personalizada" & vbCrLf & _
            "Nao = nao perguntar mais para este codigo" & vbCrLf & _
            "Cancelar = pular agora", vbYesNoCancel + vbQuestion, "Ponta personalizada")

        If resp = vbYes Then
            descricao = Trim(InputBox( _
                "Codigo do cliente: " & info(0) & vbCrLf & _
                "Descricao: " & Left(info(1), 200) & vbCrLf & vbCrLf & _
                "Informe a descricao da ponta personalizada:", "Ponta personalizada"))
            If descricao <> "" Then GravarCadastro wsCad, dictCad, CStr(chavePend), CStr(info(0)), descricao, False
        ElseIf resp = vbNo Then
            GravarCadastro wsCad, dictCad, CStr(chavePend), CStr(info(0)), "", True
        End If
    Next chavePend

    ' --- FASE 2: acrescenta a descricao ao fim da coluna F (codigos cadastrados) ---
    Application.ScreenUpdating = False
    For i = 1 To n
        chave = UCase(Trim(TextoCelula(dados(i, 1))))
        If chave <> "" Then
            If dictCad.Exists(chave) Then
                info = dictCad(chave)
                desc = info(1)
                If desc <> "" Then
                    textoF = TextoCelula(dados(i, 3))
                    If Right(textoF, Len(desc)) <> desc Then
                        With relatorio_ws.Cells(i + 3, 6)
                            .NumberFormat = "@"
                            .Value = Trim(textoF & " " & desc)
                        End With
                    End If
                End If
            End If
        End If
    Next i
    Application.ScreenUpdating = True

    If qtdSemCodigo > 0 Then
        MsgBox qtdSemCodigo & " linha(s) com item de ponta personalizada estao sem codigo do cliente (coluna D) e foram ignoradas." & _
            vbCrLf & "Linhas: " & Trim(linhasSemCodigo) & IIf(qtdSemCodigo > MAX_LINHAS_AVISO, " ...", ""), _
            vbExclamation, "Ponta personalizada"
    End If
End Sub

Private Function ContemGatilho(ByVal texto As String, ByRef gatilhos As Variant) As Boolean
    Dim i As Long
    For i = LBound(gatilhos) To UBound(gatilhos)
        If InStr(1, texto, gatilhos(i), vbTextCompare) > 0 Then
            ContemGatilho = True
            Exit Function
        End If
    Next i
End Function

Private Function TextoCelula(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    TextoCelula = CStr(v)
End Function

Private Function ObterAba(ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(ws.Name) = nome Then
            Set ObterAba = ws
            Exit Function
        End If
    Next ws
End Function

' Cria a aba de cadastro (com cabecalho) se ainda nao existir, sem trocar
' a aba ativa (o GerarRelatorio seleciona A4 do RELATORIO logo depois).
Private Function AbaCadastro() As Worksheet
    Dim ws As Worksheet, ativa As Object
    Set ws = ObterAba(ABA_CADASTRO)
    If ws Is Nothing Then
        Set ativa = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = ABA_CADASTRO
        ws.Columns("A:C").NumberFormat = "@"
        ws.Range("A1:C1").Value = Array(CAB_CODIGO, CAB_DESC, CAB_IGNORAR)
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns("A:C").AutoFit
        ativa.Activate
    End If
    Set AbaCadastro = ws
End Function

Private Function AbaItens() As Worksheet
    Dim ws As Worksheet, ativa As Object
    Set ws = ObterAba(ABA_ITENS)
    If ws Is Nothing Then
        Set ativa = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = ABA_ITENS
        ws.Columns("A:A").NumberFormat = "@"
        ws.Cells(1, 1).Value = CAB_ITENS
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(2, 1).Value = ITEM_EXEMPLO
        ws.Columns("A:A").AutoFit
        ativa.Activate
    End If
    Set AbaItens = ws
End Function

' Itens-gatilho (maiusculos). Celulas vazias e o cabecalho sao ignorados:
' InStr com texto vazio retorna 1 e dispararia a janela em todas as linhas.
Private Function CarregarGatilhos() As Object
    Dim d As Object, ws As Worksheet, arr As Variant
    Dim i As Long, ult As Long, t As String

    Set d = CreateObject("Scripting.Dictionary")
    Set ws = AbaItens()
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    arr = ws.Range("A1:B" & ult).Value
    For i = 1 To UBound(arr, 1)
        t = UCase(Trim(TextoCelula(arr(i, 1))))
        If t <> "" And t <> CAB_ITENS Then
            If Not d.Exists(t) Then d.Add t, True
        End If
    Next i
    Set CarregarGatilhos = d
End Function

' chave = codigo em maiusculas; valor = Array(linha, descricao, ignorar)
Private Function CarregarCadastro(ByVal ws As Worksheet) As Object
    Dim d As Object, arr As Variant
    Dim i As Long, ult As Long, chave As String

    Set d = CreateObject("Scripting.Dictionary")
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    arr = ws.Range("A1:C" & ult).Value
    For i = 1 To UBound(arr, 1)
        chave = UCase(Trim(TextoCelula(arr(i, 1))))
        If chave <> "" And chave <> CAB_CODIGO Then
            If Not d.Exists(chave) Then
                d.Add chave, Array(i, Trim(TextoCelula(arr(i, 2))), (UCase(Trim(TextoCelula(arr(i, 3)))) = "SIM"))
            End If
        End If
    Next i
    Set CarregarCadastro = d
End Function

' Grava (ou completa a linha ja existente do codigo) e atualiza o Dictionary
Private Sub GravarCadastro(ByVal ws As Worksheet, ByRef dict As Object, ByVal chave As String, _
                           ByVal codigo As String, ByVal desc As String, ByVal ignorar As Boolean)
    Dim lin As Long, info As Variant

    If dict.Exists(chave) Then
        info = dict(chave)
        lin = info(0)
    Else
        lin = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If lin < 2 Then lin = 2
        ws.Cells(lin, 1).NumberFormat = "@"
        ws.Cells(lin, 1).Value = codigo
    End If

    ws.Cells(lin, 2).NumberFormat = "@"
    ws.Cells(lin, 2).Value = desc
    ws.Cells(lin, 3).NumberFormat = "@"
    ws.Cells(lin, 3).Value = IIf(ignorar, "SIM", "")
    dict(chave) = Array(lin, desc, ignorar)
End Sub
