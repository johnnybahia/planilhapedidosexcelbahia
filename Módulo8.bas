Attribute VB_Name = "Módulo8"
Option Explicit

' ==========================================================================
' Cadastro perene de códigos DILLY (aba "CÓDIGOS DILLY").
' Regra observada: COD CLIENTE = CÓDIGO BASE & "-" & TAMANHO (sem unidade
' CM/MM). CÓDIGO BASE e COD ITEM são fixos por DESCRIÇÃO. Como o XML da
' Dilly deixou de trazer o "OC:" na observação, a coluna D do RELATORIO
' não pode mais ser deduzida daquele texto - passa a ser deduzida deste
' cadastro, alimentado a partir da própria DESCRIÇÃO (coluna F).
' ==========================================================================

Private Const ABA_CODIGOS As String = "CÓDIGOS DILLY"

' Garante que a aba de cadastro tenha o cabeçalho esperado. Se a aba
' estiver com o layout antigo (ou vazia), ela é limpa e recriada do zero.
Private Sub GarantirCabecalhoCodigosDilly(ws As Worksheet)
    If UCase(Trim(CStr(ws.Cells(1, 1).Value))) <> "DESCRIÇÃO" Then
        ws.Cells.Clear
        ws.Cells(1, 1).Value = "DESCRIÇÃO"
        ws.Cells(1, 2).Value = "CÓDIGO BASE"
        ws.Range("A1:B1").Font.Bold = True
        ws.Columns("A:B").AutoFit
    End If
End Sub

' Normaliza a descrição para uso como chave de comparação (evita
' duplicidade por espaço extra / maiúscula-minúscula).
Public Function NormalizarDescricao(ByVal s As Variant) As String
    Dim t As String
    t = UCase(Trim(CStr(s)))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizarDescricao = t
End Function

' Remove o sufixo de unidade (CM/MM) do tamanho, preservando o restante
' do conteúdo tal como está (ex.: "P", "M", "ÚNICO" não são alterados).
Public Function ExtrairTamanhoSemUnidade(ByVal v As Variant) As String
    Dim t As String
    t = UCase(Trim(CStr(v)))
    If Right(t, 2) = "CM" Or Right(t, 2) = "MM" Then
        t = Left(t, Len(t) - 2)
    End If
    ExtrairTamanhoSemUnidade = Trim(t)
End Function

' Carrega o cadastro atual em um Dictionary (chave = descrição normalizada,
' valor = código base). Usado para evitar Find/loop O(n²) sobre milhares
' de linhas do RELATORIO.
Public Function CarregarDictCodigosDilly() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastR As Long, i As Long
    Dim desc As String, base As String

    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(ABA_CODIGOS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set CarregarDictCodigosDilly = dict
        Exit Function
    End If

    GarantirCabecalhoCodigosDilly ws

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastR
        desc = NormalizarDescricao(ws.Cells(i, 1).Value)
        base = Trim(CStr(ws.Cells(i, 2).Value))
        If desc <> "" And base <> "" Then
            If Not dict.Exists(desc) Then dict.Add desc, base
        End If
    Next i

    Set CarregarDictCodigosDilly = dict
End Function

' Grava um novo par (descrição, código base) na aba de cadastro.
Private Sub GravarCodigoDilly(ByVal descricaoOriginal As String, ByVal baseCodigo As String)
    Dim ws As Worksheet
    Dim novaLinha As Long

    Set ws = ThisWorkbook.Sheets(ABA_CODIGOS)
    GarantirCabecalhoCodigosDilly ws

    novaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If novaLinha < 2 Then novaLinha = 2

    ws.Cells(novaLinha, 1).Value = descricaoOriginal
    ws.Cells(novaLinha, 2).NumberFormat = "@"
    ws.Cells(novaLinha, 2).Value = baseCodigo
End Sub

' --------------------------------------------------------------------
' FUNÇÃO 1 - Sincroniza o cadastro a partir do RELATORIO atualmente na
' tela. Só aproveita linhas DILLY cuja coluna D já esteja no formato
' "base-tamanho". Descrição já cadastrada com base diferente NUNCA é
' sobrescrita (cadastro antigo vence) - fica listada como conflito.
' --------------------------------------------------------------------
Public Sub SincronizarCodigosDilly()
    Dim relatorio_ws As Worksheet
    Dim dict As Object
    Dim last_row As Long, r As Long
    Dim nomeCliente As String, codCliente As String
    Dim descricaoOriginal As String, descricaoChave As String
    Dim partes() As String, baseExtraido As String
    Dim novos As Long, conflitos As String
    Dim msg As String

    On Error GoTo Falha
    Set relatorio_ws = ThisWorkbook.Sheets("RELATORIO")
    Set dict = CarregarDictCodigosDilly()

    last_row = relatorio_ws.Cells(relatorio_ws.Rows.Count, 1).End(xlUp).Row
    If last_row < 4 Then
        MsgBox "Não há dados no RELATORIO para sincronizar.", vbExclamation
        Exit Sub
    End If

    novos = 0
    conflitos = ""

    For r = 4 To last_row
        nomeCliente = UCase(Trim(relatorio_ws.Cells(r, 2).Value))
        If InStr(nomeCliente, "DILLY") > 0 Then
            codCliente = Trim(CStr(relatorio_ws.Cells(r, 4).Value))
            descricaoOriginal = Trim(CStr(relatorio_ws.Cells(r, 6).Value))
            descricaoChave = NormalizarDescricao(descricaoOriginal)

            If codCliente <> "" And descricaoChave <> "" And InStr(codCliente, "-") > 0 Then
                partes = Split(codCliente, "-")
                baseExtraido = Trim(partes(0))

                If baseExtraido <> "" Then
                    If Not dict.Exists(descricaoChave) Then
                        dict.Add descricaoChave, baseExtraido
                        GravarCodigoDilly descricaoOriginal, baseExtraido
                        novos = novos + 1
                    ElseIf dict(descricaoChave) <> baseExtraido Then
                        If InStr(conflitos, descricaoOriginal) = 0 Then
                            conflitos = conflitos & "- " & descricaoOriginal & " (cadastrado: " & _
                                dict(descricaoChave) & " / linha atual: " & baseExtraido & ")" & vbCrLf
                        End If
                    End If
                End If
            End If
        End If
    Next r

    msg = novos & " novo(s) código(s) cadastrado(s) em '" & ABA_CODIGOS & "'."
    If conflitos <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Conflitos encontrados (cadastro antigo mantido, revisar manualmente):" & _
            vbCrLf & conflitos
    End If
    MsgBox msg, vbInformation
    Exit Sub

Falha:
    MsgBox "Erro ao sincronizar códigos Dilly: " & Err.Description, vbCritical
End Sub

' --------------------------------------------------------------------
' FUNÇÃO 2 - Chamada pelo GerarRelatorio (Módulo3) para resolver, via
' cadastro, as linhas DILLY cuja descrição não tinha código conhecido.
' Pergunta UMA vez por descrição nova (não por linha), grava a resposta
' no cadastro e aplica em todas as linhas daquela descrição.
' --------------------------------------------------------------------
Public Sub ResolverPendenciasDilly(ByRef relatorio_ws As Worksheet, ByVal last_row As Long, _
                                    ByRef dictCodigos As Object, ByRef dictPendentes As Object)
    Dim chave As Variant
    Dim descricaoOriginal As String, resposta As String
    Dim naoCadastradas As String
    Dim r As Long, descChave As String, tamanho As String

    For Each chave In dictPendentes.Keys
        descricaoOriginal = dictPendentes(chave)
        resposta = Trim(InputBox( _
            "Cliente DILLY - descrição sem código cadastrado." & vbCrLf & vbCrLf & _
            "Descrição: " & descricaoOriginal & vbCrLf & vbCrLf & _
            "Informe o CÓDIGO BASE (sem o tamanho, ele é aplicado automaticamente):", _
            "Novo código Dilly"))

        If resposta <> "" Then
            If Not dictCodigos.Exists(chave) Then dictCodigos.Add chave, resposta
            GravarCodigoDilly descricaoOriginal, resposta
        Else
            naoCadastradas = naoCadastradas & "- " & descricaoOriginal & vbCrLf
        End If
    Next chave

    ' Aplica à coluna D todas as linhas DILLY que ficaram pendentes e cuja
    ' descrição foi resolvida nesta rodada.
    For r = 4 To last_row
        If InStr(UCase(Trim(relatorio_ws.Cells(r, 2).Value)), "DILLY") > 0 Then
            If Trim(CStr(relatorio_ws.Cells(r, 4).Value)) = "" Then
                descChave = NormalizarDescricao(relatorio_ws.Cells(r, 6).Value)
                If dictCodigos.Exists(descChave) Then
                    tamanho = ExtrairTamanhoSemUnidade(relatorio_ws.Cells(r, 7).Value)
                    relatorio_ws.Cells(r, 4).NumberFormat = "@"
                    relatorio_ws.Cells(r, 4).Value = dictCodigos(descChave) & "-" & tamanho
                End If
            End If
        End If
    Next r

    If naoCadastradas <> "" Then
        MsgBox "As descrições abaixo ficaram sem CÓDIGO CLIENTE (não informado):" & vbCrLf & vbCrLf & _
            naoCadastradas, vbExclamation
    End If
End Sub
